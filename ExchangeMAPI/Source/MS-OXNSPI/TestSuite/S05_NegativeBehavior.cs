namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class contains all the test cases designed to test the negative server behavior for each NSPI call.
    /// </summary>
    [TestClass]
    public class S05_NegativeBehavior : TestSuiteBase
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
        /// This test case is designed to verify the requirements related to setting the flag in NspiBind method to fAnonymousLogin.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC01_BindFlagWithfAnonymousLogin()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server. The input parameter dwFlags in this request is set to a value that is not "fAnonymousLogin".
            STAT stat = new STAT();
            stat.InitiateStat();
            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };
            FlatUID_r? serverGuid = guid;
            uint flags = 0x0;

            // A value of "fAnonymousLogin" in the input parameter dwFlags indicates that the server 
            // does not validate that the client is an authenticated user. The server MAY ignore this request.
            ErrorCodeValue result1 = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion

            #region Call NspiBind to initiate another session between the client and the server. The input parameter dwFlags in this request is set to "fAnonymousLogin".
            // Set dwFlags to fAnonymousLogin: 0x00000020.
            flags = (uint)NspiBindFlag.fAnonymousLogin;
            ErrorCodeValue result2 = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);

            #region Verify the requirements about NspiBind operation.
            if (Common.IsRequirementEnabled(1613, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1613, the value of result1 is {0}, the value of result2 is {1}", result1, result2);

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1613
                Site.CaptureRequirementIfIsTrue(
                    result1 != result2,
                    1613,
                    @"[In Appendix B: Product Behavior] Implementation does not ignore this request [a value of ""fAnonymousLogin"" in the input parameter dwFlags]. (Exchange 2010 and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R927");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R927
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.NotSupported,
                result2,
                "MS-OXCDATA",
                927,
                @"[In Error Codes] NotSupported (MAPI_E_NO_SUPPORT, ecNotSupported, ecNotImplemented) will be returned, if the server does not support this method call.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R928");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R928
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                (uint)result2,
                "MS-OXCDATA",
                928,
                @"[In Error Codes] The numeric value (hex) for error code NotSupported is 0x80040102, %x02.01.04.80.");

            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to MS-OXNSPI operations with CP_WINUNICODE CodePage.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC02_OperationsWithUnicodeCodePage()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server. The CodePage field of the input parameter pStat in this request is set to “CP_WINUNICODE”.
            STAT stat = new STAT();
            stat.InitiateStat();
            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };
            FlatUID_r? serverGuid = guid;

            uint flags = 0x0;
            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R680, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R680
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                680,
                @"[In NspiBind] [Server Processing Rules: Upon receiving message NspiBind, the server MUST process the data from the message subject to the following constraints:] [constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R696, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R696
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidCodepage,
                this.Result,
                696,
                @"[In NspiBind] [Server Processing Rules: Upon receiving message NspiBind, the server MUST process the data from the message subject to the following constraints:] [constraint 5] If the server will not service connections using that code page, the server MUST return the error code ""InvalidCodepage"".");

            #endregion

            #endregion

            #region Call NspiBind to initiate a session between the client and the server.

            // Previous Bind operation should be failed, the following Bind operation may fail several times for same reason, so need retry.
            int tryTimes = Convert.ToInt32(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            stat.InitiateStat();
            for (int i = 0; i < tryTimes; i++)
            {
                this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
                if (this.Result == ErrorCodeValue.Success)
                {
                    break;
                }
            }

            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiGetSpecialTable with the CodePage field of the input parameter pStat in this request is set to “CP_WINUNICODE”.
            uint version = 0;
            PropertyRowSet_r? rows;

            // Set flags not to containing the flag NspiUnicodeStrings (0x04) and NspiAddressCreationTemplates (0x02).
            uint flagsOfGetSpecialTable = 0;

            // Set CodePage field of pStat to containing the value CP_WINUNICODE
            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows, false);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R732, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R732
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                732,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the input parameter dwFlags does not contain the value ""NspiUnicodeStrings"", and the input parameter dwFlags does not contain the value ""NspiAddressCreationTemplates"", and the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion
            #endregion

            #region Call NspiGetSpecialTable with the CodePage field of the input parameter pStat in this request is set to an invalid code page.
            stat.CodePage = 0xff;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows, false);
            Site.Assert.AreNotEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable should not return Success!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R736");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R736
            Site.CaptureRequirementIfIsNull(
                rows,
                736,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the server returns any return value other than ""Success"", the server MUST return a NULL for the output parameter ppRows.");
            #endregion

            #region Call NspiUpdateStat with CP_WINUNICODE CodePage.
            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            uint reserved = 0;
            int? delta = 0;
            stat.Delta = 2;
            ErrorCodeValue result1 = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta, false);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R780, the value of the result1 is {0}", result1);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R780
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), result1),
                780,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion
            #endregion

            #region Call NspiGetProps with CP_WINUNICODE CodePage.

            PropertyTagArray_r? propTags;
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                PropertyTagArray_r? prop = null;
                propTags = prop;
            }
            else
            {
                PropertyTagArray_r prop = new PropertyTagArray_r
                {
                    CValues = 2
                };
                prop.AulPropTag = new uint[prop.CValues];
                prop.AulPropTag[0] = (uint)AulProp.PidTagEntryId;
                prop.AulPropTag[1] = (uint)AulProp.PidTagDisplayName;
                propTags = prop;
            }

            uint flagsOfGetProps = (uint)RetrievePropertyFlag.fEphID;
            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            PropertyRow_r? row;
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out row, false);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R873, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R873
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                873,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the CodePage field of the input parameter pStat is set to the value CP_WINUNICODE and the type of the proptags in the input parameter pPropTags is PtypString8, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion
            #endregion

            #region Call NspiGetProps with invalid CodePage.
            stat.CodePage = 0xff;
            ErrorCodeValue invalidCodePageNspiGetProps = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out row, false);
            Site.Assert.AreNotEqual<ErrorCodeValue>(ErrorCodeValue.Success, invalidCodePageNspiGetProps, "NspiGetProps should not return Success!");
            Site.Assert.AreNotEqual<ErrorCodeValue>(ErrorCodeValue.ErrorsReturned, invalidCodePageNspiGetProps, "NspiGetProps should not return ErrorsReturned!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R877");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R877
            Site.CaptureRequirementIfIsNull(
                row,
                877,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the server returns any return values other than""ErrorsReturned"" (0x00040380) or ""Success"" (0x00000000), the server MUST return a NULL for the output parameter ppRows.");

            #endregion

            #region Call NspiQueryRows with CP_WINUNICODE CodePage.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            uint tableCount = 0;
            uint[] table = null;
            uint count = 10;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagDisplayName,
                }
            };
            propTags = propTagsInstance;

            // Set the CodePage field of the input parameter pStat to containing the value CP_WINUNICODE.
            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R935, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R935
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                935,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion Capture
            #endregion

            #region Call NspiQueryRows with invalid CodePage.
            stat.CodePage = 0xff;

            // Here use the same input parameter to call NspiQueryRows.
            ErrorCodeValue invalidCodePageNspiQueryRows = this.ProtocolAdatper.NspiQueryRows(flagsOfGetProps, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows, false);
            Site.Assert.AreNotEqual<ErrorCodeValue>(ErrorCodeValue.Success, invalidCodePageNspiQueryRows, "NspiGetProps should not return Success!");
            Site.Assert.AreNotEqual<ErrorCodeValue>(ErrorCodeValue.ErrorsReturned, invalidCodePageNspiQueryRows, "NspiGetProps should not return ErrorsReturned!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R982");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R982
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                invalidCodePageNspiGetProps,
                invalidCodePageNspiQueryRows,
                982,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] If a call to the NspiGetProps method with these parameters [hRpc, dwFlags, pStat and pPropTags] would return any value other than ""Success"" or ""ErrorsReturned"", the server MUST return that error code as the return value for the NspiQueryRows method.");

            #endregion

            #region Call NspiSeekEntries with CP_WINUNICODE CodePage.
            uint reservedOfSeekEntries = 0;
            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = 0x3001001F,
                Reserved = (uint)0x00
            };

            // Set property PidTagDisplayName (0x3001) with PtypString type (0x001F).
            string displayName;
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            }
            else
            {
                displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "\0";
            }

            target.Value.LpszW = System.Text.Encoding.Unicode.GetBytes(displayName);

            PropertyTagArray_r tags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    target.PropTag
                }
            };
            PropertyTagArray_r? propTagsOfSeekEntries = tags;

            PropertyTagArray_r? tableOfSeekEntries = null;
            PropertyRowSet_r? rowsOfSeekEntries;

            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1016, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1016
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1016,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion
            #endregion

            #region Call NspiGetMatches with CP_WINUNICODE CodePage.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            Restriction_r res_r = new Restriction_r
            {
                Rt = 0x8,
                Res =
                    new RestrictionUnion_r
                    {
                        ResExist =
                            new ExistRestriction_r
                            {
                                Reserved1 = 0,
                                Reserved2 = 0,
                                PropTag = (uint)AulProp.PidTagEntryId
                            }
                    }
            };
            Restriction_r? filter = res_r;

            propTagsInstance = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTagsInstance;
            PropertyName_r? propNameOfGetMatches = null;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1101, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1101
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1101,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion
            #endregion

            #region Call NspiResortRestriction with CP_WINUNICODE CodePage.
            #region NspiDNToMId
            reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                CValues = 3,
                LppszA = new string[3]
                {
                    Common.GetConfigurationPropertyValue("User1Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User2Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User3Essdn", this.Site),
                }
            };

            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            #endregion

            #region NspiResortRestriction

            uint reservedOfResortRestriction = 0;
            PropertyTagArray_r inmids = new PropertyTagArray_r();
            inmids = mids.Value;

            // If the object specified by the CurrentRec field of the input parameter pStat is in the constructed Explicit Table, the NumPos field of the output parameter pStat is set to the numeric position in the Explicit Table.
            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName;
            outMIds = null;
            this.Result = this.ProtocolAdatper.NspiResortRestriction(reservedOfResortRestriction, ref stat, inmids, ref outMIds, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1181, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1181
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1181,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion
            #endregion
            #endregion

            #region Call NspiCompareMIds with CP_WINUNICODE CodePage.
            #region NspiDNToMId
            reserved = 0;
            names = new StringsArray_r
            {
                CValues = 2,
                LppszA = new string[2]
            };
            names.LppszA[0] = Common.GetConfigurationPropertyValue("User3Essdn", this.Site);
            names.LppszA[1] = Common.GetConfigurationPropertyValue("User1Essdn", this.Site);
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiDnToMId should return Success!");
            #endregion

            #region NspiCompareMIds
            uint mid1 = mids.Value.AulPropTag[0];
            uint mid2 = mids.Value.AulPropTag[1];
            uint reservedOfCompareMIds = 0;
            int results;

            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            this.Result = this.ProtocolAdatper.NspiCompareMIds(reservedOfCompareMIds, stat, mid1, mid2, out results, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1223, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1223
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1223,
                @"[In NspiCompareMIds] [Server Processing Rules: Upon receiving message NspiCompareMIds, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion
            #endregion
            #endregion

            #region Call NspiModProps with CP_WINUNICODE CodePage.
            #region NspiGetMatches
            // Call NspiGetMatches to get valid MIDs and rows.
            reserved1 = 0;
            reserver2 = 0;
            proReserved = null;
            requested = Constants.GetMatchesRequestedRowNumber;

            res_r = new Restriction_r
            {
                Rt = 0x8,
                Res =
                    new RestrictionUnion_r
                    {
                        ResExist =
                            new ExistRestriction_r
                            {
                                Reserved1 = 0,
                                Reserved2 = 0,
                                PropTag = (uint)AulProp.PidTagDisplayName
                            }
                    }
            };
            filter = res_r;

            propNameOfGetMatches = null;
            propTagsInstance = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagDisplayName,
                }
            };
            propTagsOfGetMatches = propTagsInstance;

            stat.CodePage = (uint)RequiredCodePage.CP_TELETEX;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            #endregion

            #region NspiModProps
            // Call NspiModProps method with specific PidTagAddressBookX509Certificate tag.
            uint reservedOfModProps = 0;
            BinaryArray_r certificate = new BinaryArray_r();
            PropertyRow_r rowOfModProps = new PropertyRow_r
            {
                LpProps = new PropertyValue_r[1]
            };
            rowOfModProps.LpProps[0].PropTag = (uint)AulProp.PidTagAddressBookX509Certificate;
            rowOfModProps.LpProps[0].Value.MVbin = certificate;

            PropertyTagArray_r instanceOfModProps = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagAddressBookX509Certificate
                }
            };
            PropertyTagArray_r? propTagsOfModProps = instanceOfModProps;

            // Get user name.
            string userName = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            // Set the CurrentRec field with the minimal entry ID of mail user name.
            for (int i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                string name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[0].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(userName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    stat.CurrentRec = outMIds.Value.AulPropTag[i];
                    break;
                }
            }

            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            this.Result = this.ProtocolAdatper.NspiModProps(reservedOfModProps, stat, propTagsOfModProps, rowOfModProps, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1280, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1280
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1280,
                @"[In NspiModProps] [Server Processing Rules: Upon receiving message NspiModProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion

            #endregion
            #endregion

            #region Call NspiResolveNames with CP_WINUNICODE CodePage.
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                StringsArray_r strArray;
                strArray.CValues = 2;
                strArray.LppszA = new string[strArray.CValues];
                strArray.LppszA[0] = Common.GetConfigurationPropertyValue("User1Name", this.Site);
                strArray.LppszA[1] = string.Empty;

                propTags = null;
                PropertyRowSet_r? rowOfResolveNames;

                stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
                this.Result = this.ProtocolAdatper.NspiResolveNames((uint)0, stat, propTags, strArray, out mids, out rowOfResolveNames, false);

                #region Capture
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1359, the value of the result is {0}", this.Result);

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1359
                Site.CaptureRequirementIfIsTrue(
                    Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                    1359,
                    @"[In NspiResolveNames] [Server Processing Rules: Upon receiving message NspiResolveNames, the server MUST process the data from the message subject to the following constraints:] 
                [Constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, 
                UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark,
                AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.2.");

                #endregion
            }
            #endregion

            #region Call NspiResolveNamesW with CP_WINUNICODE CodePage.
            uint reservedOfResolveNamesW = 0;
            WStringsArray_r wstrArray;
            wstrArray.CValues = 2;
            wstrArray.LppszW = new string[wstrArray.CValues];
            wstrArray.LppszW[0] = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            wstrArray.LppszW[1] = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            propTags = null;
            PropertyRowSet_r? rowOfResolveNamesW;

            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            this.Result = this.ProtocolAdatper.NspiResolveNamesW(reservedOfResolveNamesW, stat, propTags, wstrArray, out mids, out rowOfResolveNamesW, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1410, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1410
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1410,
                @"[In NspiResolveNamesW] [Server Processing Rules: Upon receiving message NspiResolveNamesW, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the CodePage field of the input parameter pStat contains the value CP_WINUNICODE, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiUpdateStat returning InvalidBookmark.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC03_UpdateStatFailedWithInvalidBookmark()
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

            #region Call NspiUpdateStat with an invalid ContainerID field in the input parameter pStat, so that the address book container specified by the invalid ContainerID field cannot be located.
            uint reserved = 0;
            int? delta = 1;
            stat.ContainerID = (uint)MinimalEntryID.MID_CURRENT;

            // statSave is used to save STAT before calling NspiUpdateStat.
            STAT statSave = stat;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R789");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R789
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidBookmark,
                this.Result,
                789,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] If the server is unable to locate the address book container specified by the ContainerID field in the input parameter pStat, the server MUST return the return value InvalidBookmark.");

            this.VerifyWhetherpStatIsModifiedForNspiUpdateStat(statSave, stat);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R991");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R991
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidBookmark,
                this.Result,
                "MS-OXCDATA",
                991,
                @"[In Error Codes] InvalidBookmark(MAPI_E_INVALID_BOOKMARK, ecInvalidBookmark) will be returned, if the bookmark passed to
                a table operation was not created on the same table.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R992");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R992
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040405,
                (uint)this.Result,
                "MS-OXCDATA",
                992,
                @"[In Error Codes] The numeric value (hex) for error code InvalidBookmark is 0x80040405, %x05.04.04.80.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiUpdateStat returning NotFound.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC04_UpdateStatFailedWithNotFound()
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

            #region Call NspiUpdateStat with an invalid CurrentRec field in the input parameter pStat, so that the row specified by the invalid CurrentRec field cannot be found.
            uint reserved = 0;
            int? delta = 1;

            // Set CurrentRec field of the input parameter STAT with an invalid value.
            stat.CurrentRec = uint.Parse(Constants.UnrecognizedMID);

            // statSave is used to save STAT before calling NspiUpdateStat.
            STAT statSave = stat;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);

            #region Capture code

            this.VerifyWhetherpStatIsModifiedForNspiUpdateStat(statSave, stat);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R792");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R792
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.NotFound,
                this.Result,
                792,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 5: The server locates the initial position row in the table specified by the ContainerID field of the input parameter pStat as follows:] If the row [the row specified by the CurrentRec field of the input parameter pStat] cannot be found, the server MUST return the error ""NotFound"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R949");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R949
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.NotFound,
                this.Result,
                "MS-OXCDATA",
                949,
                @"[In Error Codes] NotFound(MAPI_E_NOT_FOUND, ecNotFound, ecAttachNotFound, ecUnknownRecip, ecPropNotExistent) will be returned,
                if the requested object could not be found at the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R950");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R950
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                (uint)this.Result,
                "MS-OXCDATA",
                950,
                @"[In Error Codes] The numeric value (hex) for error code NotFound is 0x8004010F, %x0F.01.04.80.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiUpdateStat returning ErrorsReturned.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC05_GetPropsFailedWithErrorsReturned()
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

            #region Call NspiUpdateStat to update the positioning changes in a table.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiUpdateStat to get a property that does not exist in a specified table row.
            PropertyTagArray_r prop = new PropertyTagArray_r
            {
                CValues = 1
            };
            prop.AulPropTag = new uint[prop.CValues];

            // PtypString8 type value: 0x0000001E.
            prop.AulPropTag[0] = (uint)PropertyTypeValue.PtypString8;
            PropertyTagArray_r? propTags = prop;
            uint flagsOfGetProps = (uint)RetrievePropertyFlag.fEphID;
            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            PropertyRow_r? rows;

            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R909");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R909
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.ErrorsReturned,
                this.Result,
                909,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 12] If a property in the proptag list has no value on the object specified by the CurrentRec field, the server MUST return the error code ErrorsReturned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R910");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R910
            // The last four bytes of the proptag field represents the type of the property. If its value is 0x000a, it means the property has no value.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x000a,
                rows.Value.LpProps[0].PropTag & 0x0000ffff,
                910,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 12] [If a property in the proptag list has no value on the object specified by the CurrentRec field] The server MUST set the aulPropTag member corresponding to the proptag with no value with the proptag that has no value with the PtypErrorCode property type.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2070");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2070
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.ErrorsReturned,
                this.Result,
                "MS-OXCDATA",
                2070,
                @"[In Warning Codes] ErrorsReturned (MAPI_W_ERRORS_RETURNED, ecWarnWithErrors) will be returned, if a request involving multiple properties failed for one or more individual properties, while succeeding overall.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2071");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2071
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00040380,
                (uint)this.Result,
                "MS-OXCDATA",
                2071,
                @"[In Warning Codes] The numeric value (hex) for error code ErrorsReturned is 0x00040380, %x80.03.04.00.");

            #endregion Capture
            #endregion

            #region Call NspiGetProps method with CurrentRec field of stat set to the value that server can't locate.
            stat.CurrentRec = (uint)MinimalEntryID.MID_END_OF_TABLE;
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.ErrorsReturned, this.Result, "NspiGetProps should not return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R907");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R907
            Site.CaptureRequirementIfAreEqual<uint>(
                0xa,
                rows.Value.LpProps[0].PropTag & 0x0000ffff,
                907,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 11] If the server is unable to locate the object specified in the CurrentRec field of the input parameter pStat, the server MUST proceed as if the object was located but had no values for any properties.");

            #endregion Capture

            #endregion

            #region Call NspiUpdateStat with the CurrentRec field set to MID_CURRENT.
            prop.AulPropTag[0] = (uint)AulProp.PidTagDisplayName;
            propTags = prop;
            stat.CurrentRec = (uint)MinimalEntryID.MID_CURRENT;
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1985");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1985
            Site.CaptureRequirementIfAreNotEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1985,
                @"[In Positioning Minimal Entry IDs] [MID_CURRENT] For method NspiGetProps, it is an invalid Minimal Entry ID, guaranteed to not specify any object in the address book.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiSeekEntries returning InvalidBookmark.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC06_SeekEntriesFailedWithInvalidBookmark()
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

            #region Call NspiSeekEntries with an invalid ContainerID field in the input parameter pStat, so that the address book container specified by the invalid ContainerID field cannot be located.
            uint reservedOfSeekEntries = 0;

            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = (uint)0x00
            };
            string displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName + "\0");
            }

            PropertyTagArray_r tags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    target.PropTag
                }
            };
            PropertyTagArray_r? propTagsOfSeekEntries = tags;

            PropertyTagArray_r? tableOfSeekEntries = null;
            PropertyRowSet_r? rowsOfSeekEntries;

            stat.ContainerID = (uint)MinimalEntryID.MID_CURRENT;
            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1033");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1033
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidBookmark,
                this.Result,
                1033,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] If the server is unable to locate the address book container specified by the ContainerID field in the input parameter pStat, the server MUST return the return value ""InvalidBookmark"".");

            this.VerifyParametersRelatedWithNspiSeekEntries(this.Result, rowsOfSeekEntries, inputStat, stat);

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiSeekEntries returning NotFound.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC07_SeekEntriesFailedWithNotFound()
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

            #region Call NspiSeekEntries with an invalid input parameter pTarget, so that the row specified by the invalid pTarget cannot be found.
            // If the input parameter Reserved contains any value other than 0, the server MUST return one of the return values specified in section 2.2.2.
            uint reservedOfSeekEntries = 0;

            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = (uint)0x00,
            };
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = new byte[] { 0xff, 0xff, 0xff, 0xff };
            }
            else
            {
                target.Value.LpszA = new byte[] { 0xff, 0xff, 0xff, 0xff, 0x00 };
            }

            PropertyTagArray_r tags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    target.PropTag
                }
            };
            PropertyTagArray_r? propTagsOfSeekEntries = tags;

            PropertyTagArray_r? tableOfSeekEntries = null;
            PropertyRowSet_r? rowsOfSeekEntries;

            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries);

            #region Capture

            this.VerifyParametersRelatedWithNspiSeekEntries(this.Result, rowsOfSeekEntries, inputStat, stat);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1048");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1048
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.NotFound,
                this.Result,
                1048,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 13] If no such row [the first row in the specified table that has a value equal to or greater than the value specified in the input parameter pTarget] exists, the server MUST return the value NotFound.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiSeekEntries returning GeneralFailure.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC08_SeekEntriesFailedWithGeneralFailure()
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

            #region Call NspiSeekEntries with the SortType field of the input parameter pStat that has the value SortTypePhoneticDisplayName.
            // If the input parameter Reserved contains any value other than 0, the server MUST return one of the return values specified in section 2.2.2. 
            uint reservedOfSeekEntries = 0;

            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = (uint)0x00
            };
            string displayName;
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            }
            else
            {
                displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "\0";
            }

            target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);

            PropertyTagArray_r tags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    target.PropTag
                }
            };
            PropertyTagArray_r? propTagsOfSeekEntries = tags;

            PropertyTagArray_r? tableOfSeekEntries = null;
            PropertyRowSet_r? rowsOfSeekEntries;

            // If the server does not support the SortTypePhoneticDisplayName and the SortType field of the input parameter pStat has the value SortTypePhoneticDisplayName, the server MUST return the value GeneralFailure.
            stat.SortType = (uint)TableSortOrder.SortTypePhoneticDisplayName;
            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries, false);

            #region Capture

            this.VerifyParametersRelatedWithNspiSeekEntries(this.Result, rowsOfSeekEntries, inputStat, stat);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1039");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1039
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.GeneralFailure,
                this.Result,
                1039,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 9] If the server does not support the SortTypePhoneticDisplayName and the SortType field of the input parameter pStat has the value SortTypePhoneticDisplayName, the server MUST return the value GeneralFailure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R899");

            // This requirement is abstracted from DTD, which is a general description for many Open Specifications. In MS-OXNSPI, it returns GeneralFailure in specific conditions.
            // Verify MS-OXCDATA requirement: MS-OXCDATA_R899
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.GeneralFailure,
                this.Result,
                "MS-OXCDATA",
                899,
                @"[In Error Codes] GeneralFailure (E_FAIL, MAPI_E_CALL_FAILED, ecError, SYNC_E_ERROR) will be returned, if the operation failed for an unspecified reason.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R900");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R900
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                (uint)this.Result,
                "MS-OXCDATA",
                900,
                @"[In Error Codes] The numeric value (hex) for error code GeneralFailure is 0x80004005, %x05.40.00.80.");

            #endregion Capture
            #endregion

            #region Call NspiSeekEntries with the SortType field of the input parameter pStat that has a value that is not SortTypeDisplayName or SortTypePhoneticDisplayName.
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName_RO;
            inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries, false);

            #region Capture

            this.VerifyParametersRelatedWithNspiSeekEntries(this.Result, rowsOfSeekEntries, inputStat, stat);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1041");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1041
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.GeneralFailure,
                this.Result,
                1041,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 10] If the SortType field in the input parameter pStat has any value other than SortTypeDisplayName or SortTypePhoneticDisplayName, the server MUST return the value GeneralFailure.");

            #endregion Capture
            #endregion

            #region Call NspiSeekEntries with the SortType field in the input parameter pStat that is SortTypeDisplayName and the property specified in the input parameter pTarget that is not PidTagDisplayName.
            // If the SortType field in the input parameter pStat is SortTypeDisplayName and the property specified in the input parameter pTarget is anything other than PidTagDisplayName (with either the Property Type PtypString8 or PtypString), the server MUST return the value GeneralFailure.
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName;
            target.PropTag = (uint)AulProp.PidTagDisplayType;
            target.Reserved = 0;
            target.Value.L = 0;

            tags.CValues = 1;
            tags.AulPropTag[0] = target.PropTag;
            propTagsOfSeekEntries = tags;

            inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries, false);

            #region Capture

            this.VerifyParametersRelatedWithNspiSeekEntries(this.Result, rowsOfSeekEntries, inputStat, stat);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1043");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1043
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.GeneralFailure,
                this.Result,
                1043,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 11] If the SortType field in the input parameter pStat is SortTypeDisplayName and the property specified in the input parameter pTarget is anything other than PidTagDisplayName (with either the Property Type PtypString8 or PtypString), the server MUST return the value GeneralFailure.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiSeekEntries with the input parameter Reserved that contains any value other than 0.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC09_SeekEntriesWithReservedNonZero()
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

            #region Call NspiUpdateStat to update the STAT block that represents the position in a table to reflect positioning changes requested by the client.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiGetPropList method with dwFlags set to fEphID and CodePage field of stat not set to CP_WINUNICODE.
            uint flagsOfGetPropList = (uint)RetrievePropertyFlag.fEphID;
            PropertyTagArray_r? propTagsOfGetPropList;
            uint codePage = (uint)RequiredCodePage.CP_TELETEX;

            this.Result = this.ProtocolAdatper.NspiGetPropList(flagsOfGetPropList, stat.CurrentRec, codePage, out propTagsOfGetPropList);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetPropList should return Success!");
            #endregion

            #region Call NspiSeekEntries with input parameter Reserved set to 1.
            uint reservedOfSeekEntries = 0x1;
            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = (uint)0x00
            };
            string displayName;
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            }
            else
            {
                displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "\0";
            }

            target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);

            PropertyTagArray_r tags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    target.PropTag
                }
            };
            PropertyTagArray_r? propTagsOfSeekEntries = tags;

            PropertyTagArray_r? tableOfSeekEntries = null;
            PropertyRowSet_r? rowsOfSeekEntries;

            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1024, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1024
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1024,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the input parameter Reserved contains any value other than 0, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiQueryRows returning InvalidBookmark.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC10_QueryRowsFailedWithInvalidBookmark()
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

            #region Call NspiQueryRows with lpETable set to NULL and pStat containing an invalid ContainerID field, so that the address book container specified by the invalid ContainerID field cannot be located.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fSkipObjects;
            uint tableCount = 0;
            uint[] table = null;
            uint count = 1;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 4,
                AulPropTag = new uint[4]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            PropertyTagArray_r? propTags = propTagsInstance;

            // If the input parameter lpETable is NULL and the server is unable to locate the address book container specified by the ContainerID field in the input parameter pStat, the server MUST return the return value "InvalidBookmark".
            stat.ContainerID = (uint)MinimalEntryID.MID_CURRENT;
            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R948");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R948
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidBookmark,
                this.Result,
                948,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If the input parameter lpETable is NULL and the server is unable to locate the address book container specified by the ContainerID field in the input parameter pStat, the server MUST return the return value ""InvalidBookmark"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R943");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R943
            Site.CaptureRequirementIfIsNull(
                rowsOfQueryRows,
                943,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the server returns any return values other than ""Success"", the server MUST return a NULL for the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1762: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
                "the SortType of inputStat is {9}, the ContainerID of inputStat is {10}, the CurrentRec of inputStat is {11}, the Delta of inputStat is {12}, the NumPos of inputStat is {13}, the TotalRecs of inputStat is {13}, the CodePage of inputStat is {14}, the TemplateLocale of inputStat is {15}, the SortLocale of inputStat is {16}",
                stat.SortType,
                stat.ContainerID,
                stat.CurrentRec,
                stat.Delta,
                stat.NumPos,
                stat.TotalRecs,
                stat.CodePage,
                stat.TemplateLocale,
                stat.SortLocale,
                inputStat.SortType,
                inputStat.ContainerID,
                inputStat.CurrentRec,
                inputStat.Delta,
                inputStat.NumPos,
                inputStat.TotalRecs,
                inputStat.CodePage,
                inputStat.TemplateLocale,
                inputStat.SortLocale);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1762
            bool isVerifyR1762 = (stat.SortType == inputStat.SortType)
                                && (stat.ContainerID == inputStat.ContainerID)
                                && (stat.CurrentRec == inputStat.CurrentRec)
                                && (stat.Delta == inputStat.Delta)
                                && (stat.NumPos == inputStat.NumPos)
                                && (stat.TotalRecs == inputStat.TotalRecs)
                                && (stat.CodePage == inputStat.CodePage)
                                && (stat.TemplateLocale == inputStat.TemplateLocale)
                                && (stat.SortLocale == inputStat.SortLocale);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1762,
                1762,
                @"[In NspiQueryRows] If the server returns any return values other than ""Success"", the server MUST NOT modify
                the output parameter pStat.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiQueryRows with the input parameter lpETable set to NULL and the input parameter Count set to 0.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC11_QueryRowsETableNullCountZero()
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

            #region Call NspiQueryRows with lpETable set to NULL and Count set to 0.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            uint tableCount = 0;
            uint[] table = null;
            uint count = 0;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagDisplayName,
                }
            };

            PropertyTagArray_r? propTags = propTagsInstance;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R939, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R939
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                939,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the input parameter lpETable is NULL and the input parameter Count is 0, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiModLinkAtt returning InvalidParameter.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC12_ModLinkAttFailedWithInvalidParameter()
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

            #region Call NspiGetMatches to get an Explicit Table.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            // Create Restriction_r structure to use the display name of specific user as the filter parameter of NspiGetMatches method.
            Restriction_r propertyRestriction = new Restriction_r
            {
                Rt = 0x04,
                Res = new RestrictionUnion_r
                {
                    ResProperty = new Propertyrestriction_r
                    {
                        Relop = 0x04
                    }
                }
            };
            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = 0
            };
            string userName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(userName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(userName + "\0");
            }

            propertyRestriction.Res.ResProperty.Prop = new PropertyValue_r[] { target };
            propertyRestriction.Res.ResProperty.PropTag = (uint)AulProp.PidTagDisplayName;

            Restriction_r? filter = propertyRestriction;

            PropertyName_r? propNameOfGetMatches = null;
            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    unchecked((uint)AulProp.PidTagAddressBookPublicDelegates),
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            #endregion

            #region Call NspiModLinkAtt with an invalid input parameter dwMId, so that the server cannot locate the object specified by the invalid dwMId.
            uint flagsOfModLinkAtt = 0; // A value which does not contain fDelete flag (0x1).
            BinaryArray_r entryId = new BinaryArray_r
            {
                CValues = 1,
                Lpbin = new Binary_r[1]
            };

            // Get user name.
            for (int i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                string name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(userName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    entryId.Lpbin[0] = rowsOfGetMatches.Value.ARow[i].LpProps[0].Value.Bin;
                    break;
                }
            }

            // Modify PidTagAddressBookMember.
            uint propTagOfModLinkAtt = (uint)AulProp.PidTagAddressBookMember;

            // Set MID to an invalid value.
            uint midOfModLinkAtt = uint.Parse(Constants.UnrecognizedMID);
            this.Result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1328");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1328
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidParameter,
                this.Result,
                1328,
                @"[In NspiModLinkAtt] [Server Processing Rules: Upon receiving message NspiModLinkAtt, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] If the server is unable to locate the object specified by the input parameter dwMId, the server MUST return the value ""InvalidParameter"" (0x80070057).");

            #endregion
            #endregion

            #region Call NspiGetMatches again to see if any property is modified.
            PropertyRowSet_r? rowsOfGetMatches1;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            #region Capture code
            if (Common.IsRequirementEnabled(2003010, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R2003010");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R2003010
                Site.CaptureRequirementIfIsTrue(
                    AdapterHelper.AreTwoPropertyRowSetEqual(rowsOfGetMatches, rowsOfGetMatches1),
                    2003010,
                    @"[In Appendix A: Product Behavior] Implementation does not modify any properties of any objects in the address book. (Microsoft Exchange Server 2010 Service Pack 3 (SP3) follows this behavior).");
            }
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiModLinkAtt returning NotFound.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC13_ModLinkAttFailedWithNotFound()
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

            #region Call NspiGetMatches to get an Explicit Table.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            // Create Restriction_r structure to use the display name of specific user as the filter parameter of NspiGetMatches method.
            Restriction_r propertyRestriction = new Restriction_r
            {
                Rt = 0x04,
                Res = new RestrictionUnion_r
                {
                    ResProperty = new Propertyrestriction_r
                    {
                        Relop = 0x04
                    }
                }
            };
            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = 0
            };
            string userName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(userName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(userName + "\0");
            }

            propertyRestriction.Res.ResProperty.Prop = new PropertyValue_r[] { target };
            propertyRestriction.Res.ResProperty.PropTag = (uint)AulProp.PidTagDisplayName;

            Restriction_r? filter = propertyRestriction;

            PropertyName_r? propNameOfGetMatches = null;
            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    unchecked((uint)AulProp.PidTagAddressBookPublicDelegates),
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            #endregion

            #region Call NspiModLinkAtt with an invalid input parameter ulPropTag, so that the server cannot recognize the proptag specified by the invalid ulPropTag.
            uint flagsOfModLinkAtt = 0; // A value which does not contain fDelete flag (0x1).
            uint midOfModLinkAtt = 0;
            BinaryArray_r entryId = new BinaryArray_r
            {
                CValues = 1,
                Lpbin = new Binary_r[1]
            };

            // Get user name.
            for (int i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                string name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(userName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    midOfModLinkAtt = outMIds.Value.AulPropTag[i];
                    entryId.Lpbin[0] = rowsOfGetMatches.Value.ARow[i].LpProps[0].Value.Bin;
                    break;
                }
            }

            uint propTagOfModLinkAtt = (uint)AulProp.PidTagDisplayName;
            this.Result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1326");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1326
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.NotFound,
                this.Result,
                1326,
                @"[In NspiModLinkAtt] [Server Processing Rules: Upon receiving message NspiModLinkAtt, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the input parameter ulPropTag does not specify a proptag the server recognizes, the server MUST return NotFound.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiModLinkAtt returning AccessDenied.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC14_ModLinkAttFailedWithAccessDenied()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0; // A value which does not contain fDelete flag (0x1).
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

            #region Call NspiGetMatches to get an Explicit Table.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            // Create Restriction_r structure to use the display name of specific user as the filter parameter of NspiGetMatches method.
            Restriction_r propertyRestriction1 = new Restriction_r
            {
                Rt = 0x04,
                Res = new RestrictionUnion_r
                {
                    ResProperty = new Propertyrestriction_r
                    {
                        Relop = 0x04
                    }
                }
            };
            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = 0
            };
            string memberName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(memberName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(memberName + "\0");
            }

            propertyRestriction1.Res.ResProperty.Prop = new PropertyValue_r[] { target };
            propertyRestriction1.Res.ResProperty.PropTag = (uint)AulProp.PidTagDisplayName;

            Restriction_r propertyRestriction2 = new Restriction_r
            {
                Rt = 0x04,
                Res = new RestrictionUnion_r
                {
                    ResProperty = new Propertyrestriction_r
                    {
                        Relop = 0x04
                    }
                }
            };
            string agentName = Common.GetConfigurationPropertyValue("AgentName", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(agentName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(agentName + "\0");
            }

            propertyRestriction2.Res.ResProperty.Prop = new PropertyValue_r[] { target };
            propertyRestriction2.Res.ResProperty.PropTag = (uint)AulProp.PidTagDisplayName;

            Restriction_r restrictionOr = new Restriction_r
            {
                Rt = 0x01,
                Res =
                    new RestrictionUnion_r
                    {
                        ResOr = new OrRestriction_r
                        {
                            CRes = 2,
                            LpRes = new Restriction_r[]
                            {
                                propertyRestriction1, propertyRestriction2
                            }
                        }
                    }
            };

            Restriction_r? filter = restrictionOr;

            PropertyName_r? propNameOfGetMatches = null;
            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    unchecked((uint)AulProp.PidTagAddressBookMember),
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            #endregion

            #region Call NspiModLinkAtt to modify the value of the PidTagAddressBookMember property of an address book object with display type DT_AGENT, which is not allowed by the server.

            // A value which does not contain fDelete flag (0x1).
            uint flagsOfModLinkAtt = 0;
            uint propTagOfModLinkAtt = (uint)AulProp.PidTagAddressBookMember;
            uint midOfModLinkAtt = 0;
            BinaryArray_r entryId = new BinaryArray_r
            {
                CValues = 1,
                Lpbin = new Binary_r[1]
            };

            for (int i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                string name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(agentName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    midOfModLinkAtt = outMIds.Value.AulPropTag[i];
                }
                else if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(memberName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    entryId.Lpbin[0] = rowsOfGetMatches.Value.ARow[i].LpProps[0].Value.Bin;
                }

                if (midOfModLinkAtt != 0 && entryId.Lpbin[0].Cb != 0)
                {
                    break;
                }
            }
            
            // Modify the PidTagAddressBookMember.
            this.Result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.AccessDenied, this.Result, "NspiModLinkAtt method should return access denied.");
            #endregion

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1647");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1647
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.AccessDenied,
                this.Result,
                1647,
                @"[In NspiModLinkAtt] [Server Processing Rules: Upon receiving message NspiModLinkAtt, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] [If the server is able to locate the object, but will not allow modifications to the object due to its display type,] the server MUST return the value AccessDenied (0x80070005).");
            #endregion

            #region Call NspiGetMatches to get an Explicit Table.
            // Output parameters.
            PropertyTagArray_r? outMIds1;
            PropertyRowSet_r? rowsOfGetMatches1;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds1, propTagsOfGetMatches, out rowsOfGetMatches1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            #endregion

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1646");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1646
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowSetEqual(rowsOfGetMatches, rowsOfGetMatches1),
                1646,
                @"[In NspiModLinkAtt] [Server Processing Rules: Upon receiving message NspiModLinkAtt, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If the server is able to locate the object, but will not allow modifications to the object due to its display type, the server MUST NOT modify any properties of any objects in the address book.");
            #endregion

            #region Call NspiModLinkAtt to modify the value of the PidTagAddressBookPublicDelegates property of an address book object with display type DT_AGENT, which is not allowed by the server.
            // Set an invalid mail user name.
            string userName = Common.GetConfigurationPropertyValue("AgentName", this.Site);
            for (int i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                string name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(userName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    midOfModLinkAtt = outMIds.Value.AulPropTag[i];
                    entryId.Lpbin[0] = rowsOfGetMatches.Value.ARow[i].LpProps[0].Value.Bin;
                    break;
                }
            }

            // Modify PidTagAddressBookPublicDelegates.
            propTagOfModLinkAtt = (uint)AulProp.PidTagAddressBookPublicDelegates;
            this.Result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1338");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1338
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.AccessDenied,
                this.Result,
                1338,
                @"[In NspiModLinkAtt] [Server Processing Rules: Upon receiving message NspiModLinkAtt, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] If the server is unable to apply the modifications specified, the server MUST return the value ""AccessDenied"" (0x80070005).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R907");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R907
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.AccessDenied,
                this.Result,
                "MS-OXCDATA",
                907,
                @"[In Error Codes] AccessDenied(E_ACCESSDENIED, MAPI_E_NO_ACCESS, ecaccessdenied, ecpropsecurityviolation) will be returned, if the caller does not have sufficient access rights to perform the operation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R908");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R908
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80070005,
                (uint)this.Result,
                "MS-OXCDATA",
                908,
                @"[In Error Codes] The numeric value (hex) for error code AccessDenied is 0x80070005, %x05.00.07.80.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetMatches returning TooComplex.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC15_GetMatchesFailedWithTooComplex()
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

            #region Call NspiGetMatches with the input parameter pReserved not set to NULL.
            uint reserved1 = 0;
            uint reserver2 = 0;
            uint requested = Constants.GetMatchesRequestedRowNumber;
            PropertyName_r? propNameOfGetMatches = null;

            Restriction_r res_r = new Restriction_r
            {
                Rt = 0x8,
                Res =
                    new RestrictionUnion_r
                    {
                        ResExist = new ExistRestriction_r
                        {
                            Reserved1 = 0,
                            Reserved2 = 0,
                            PropTag = 0x0
                        }
                    }
            };
            Restriction_r? filter = res_r;

            // If the reserved input parameter pReserved contains any value other than NULL, the server MUST return the value "TooComplex".
            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagEntryId,
                }
            };
            PropertyTagArray_r? proReserved = propTags;
            PropertyTagArray_r? propTagsOfGetMatches = null;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;
            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);

            #region Capture

            this.VerifyParametersRelatedWithNspiGetMatches(this.Result, outMIds, rowsOfGetMatches, stat, inputStat);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R965");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R965
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.TooComplex,
                this.Result,
                "MS-OXCDATA",
                965,
                @"[In Error Codes] TooComplex(MAPI_E_TOO_COMPLEX, ecTooComplex) will be returned, if the operation requested is too complex for
                the server to handle; often applied to restrictions.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R966");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R966
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040117,
                (uint)this.Result,
                "MS-OXCDATA",
                966,
                @"[In Error Codes] The numeric value (hex) for error code TooComplex is 0x80040117, %x17.01.04.80.");

            #endregion capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetMatches returning InvalidBookmark.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC16_GetMatchesFailedWithInvalidBookmark()
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

            #region Call NspiGetMatches with the SortType field of the input parameter pStat set to SortTypeDisplayName and the ContainerID field of the input parameter pStat set to an invalid value.
            uint reserved1 = 0;
            uint reserver2 = 0;
            uint requested = Constants.GetMatchesRequestedRowNumber;
            PropertyName_r? propNameOfGetMatches = null;
            PropertyTagArray_r? proReserved = null;

            Restriction_r res_r = new Restriction_r
            {
                Rt = 0x8,
                Res =
                    new RestrictionUnion_r
                    {
                        ResExist =
                            new ExistRestriction_r
                            {
                                Reserved1 = 0,
                                Reserved2 = 0,
                                PropTag = (uint)AulProp.PidTagEntryId
                            }
                    }
            };
            Restriction_r? filter = res_r;

            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagEntryId,
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            // If the input SortType field of the input parameter pStat is SortTypeDisplayName or SortTypePhoneticDisplayName and the server is unable to 
            // locate the address book container specified by the ContainerID field in the input parameter pStat, the server MUST return the return value "InvalidBookmark".
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName;
            stat.ContainerID = (uint)MinimalEntryID.MID_CURRENT;
            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1122");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1122
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidBookmark,
                this.Result,
                1122,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] If the input SortType field of the input parameter pStat is SortTypeDisplayName or SortTypePhoneticDisplayName and the server is unable to locate the address book container specified by the ContainerID field in the input parameter pStat, the server MUST return the return value ""InvalidBookmark"".");

            this.VerifyParametersRelatedWithNspiGetMatches(this.Result, outMIds, rowsOfGetMatches, stat, inputStat);

            #endregion capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetMatches returning TableTooBig.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC17_GetMatchesFailedWithTableTooBig()
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

            #region Call NspiGetMatches with ulRequested set to 0 which is less than the number of rows in the Explicit Table.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;

            // Set request value less than the count of the row number.
            uint requested = 0;
            PropertyName_r? propNameOfGetMatches = null;

            Restriction_r res_r = new Restriction_r
            {
                Rt = 0x8,
                Res =
                    new RestrictionUnion_r
                    {
                        ResExist =
                            new ExistRestriction_r
                            {
                                Reserved1 = 0,
                                Reserved2 = 0,
                                PropTag = (uint)AulProp.PidTagEntryId
                            }
                    }
            };
            Restriction_r? filter = res_r;

            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagEntryId,
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;
            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1149");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1149
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.TableTooBig,
                this.Result,
                1149,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 16] If the number of rows in the constructed Explicit Table is greater than the input parameter ulRequested, the server MUST return the value ""TableTooBig"".");

            this.VerifyParametersRelatedWithNspiGetMatches(this.Result, outMIds, rowsOfGetMatches, stat, inputStat);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R989");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R989
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.TableTooBig,
                this.Result,
                "MS-OXCDATA",
                989,
                @"[In Error Codes] TableTooBig(MAPI_E_TABLE_TOO_BIG, ecTableTooBig) will be returned, if the table is too big for the requested operation to complete.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R990");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R990
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040403,
                (uint)this.Result,
                "MS-OXCDATA",
                990,
                @"[In Error Codes] The numeric value (hex) for error code TableTooBig is 0x80040403, %x03.04.04.80.");

            #endregion capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetMatches with the input parameter Filter containing any value other than NULL and the SortType field of the input parameter pStat containing any value other than SortTypeDisplayName or SortTypePhoneticDisplayName.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC18_GetMatchesWithFilterNotNull_SortOrderOtherThanDisplayNameOrPhoneticDisplayName()
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

            #region Call NspiGetMatches, the input parameter Filter is not NULL, and the SortType field of the input parameter pStat is not SortTypeDisplayName or SortTypePhoneticDisplayName.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            Restriction_r res_r = new Restriction_r
            {
                Rt = 0x8,
                Res =
                    new RestrictionUnion_r
                    {
                        ResExist =
                            new ExistRestriction_r
                            {
                                Reserved1 = 0,
                                Reserved2 = 0,
                                PropTag = (uint)AulProp.PidTagEntryId
                            }
                    }
            };
            Restriction_r? filter = res_r;

            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;
            PropertyName_r? propNameOfGetMatches = null;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            // The sort type is not SortTypeDisplayName (0x0) or SortTypePhoneticDisplayName (0x3).
            stat.SortType = 0x4;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1105, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1105
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1105,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the input parameter Filter contains any value other than NULL and the SortType field of the input parameter pStat contains any value other than SortTypeDisplayName or SortTypePhoneticDisplayName, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetMatches with the input parameter Reserved1 containing any value other than 0.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC19_GetMatchesWithNonZeroReserved1()
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

            #region Call NspiGetMatches with the input parameter Reserved1 set to 1.
            uint reserved1 = 0x1;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            Restriction_r res_r = new Restriction_r
            {
                Rt = 0x8,
                Res =
                    new RestrictionUnion_r
                    {
                        ResExist =
                            new ExistRestriction_r
                            {
                                Reserved1 = 0,
                                Reserved2 = 0,
                                PropTag = (uint)AulProp.PidTagEntryId
                            }
                    }
            };
            Restriction_r? filter = res_r;

            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;
            PropertyName_r? propNameOfGetMatches = null;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1109, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1109
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1109,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the input parameter Reserved1 contains any value other than 0, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiResortRestriction returning GeneralFailure.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC20_ResortRestrictionFailedWithGeneralFailure()
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

            #region Call NspiDNToMId to map DN to Minimal Entry ID which will be used as the input parameter in the next step.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                CValues = 3,
                LppszA = new string[3]
                {
                    Common.GetConfigurationPropertyValue("User1Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User2Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User3Essdn", this.Site),
                }
            };

            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiDNToMId should return Success!");
            #endregion

            #region Call NspiResortRestriction with the SortType field of the input parameter pStat that has the value SortTypePhoneticDisplayName, but the server does not support SortTypePhoneticDisplayName.
            uint reservedOfResortRestriction = 0;
            PropertyTagArray_r inmids = new PropertyTagArray_r();
            inmids = mids.Value;
            PropertyTagArray_r? outMIds = null;

            stat.SortType = (uint)TableSortOrder.SortTypePhoneticDisplayName;
            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiResortRestriction(reservedOfResortRestriction, ref stat, inmids, ref outMIds, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1194");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1194
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.GeneralFailure,
                this.Result,
                1194,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If the server does not support the SortTypePhoneticDisplayName and the SortType field of the input parameter pStat has the value ""SortTypePhoneticDisplayName"", the server MUST return the value ""GeneralFailure"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1189");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1189
            // When test code reaches here, server returns a value that is not "Success". So only whether outMIds is null or not needs to be determined.
            Site.CaptureRequirementIfIsNull(
                outMIds,
                1189,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the server returns any return values other than ""Success"", the server MUST return a NULL for the output parameter ppOutMIds.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1757: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
                "the SortType of inputStat is {9}, the ContainerID of inputStat is {10}, the CurrentRec of inputStat is {11}, the Delta of inputStat is {12}, the NumPos of inputStat is {13}, the TotalRecs of inputStat is {13}, the CodePage of inputStat is {14}, the TemplateLocale of inputStat is {15}, the SortLocale of inputStat is {16}",
                stat.SortType,
                stat.ContainerID,
                stat.CurrentRec,
                stat.Delta,
                stat.NumPos,
                stat.TotalRecs,
                stat.CodePage,
                stat.TemplateLocale,
                stat.SortLocale,
                inputStat.SortType,
                inputStat.ContainerID,
                inputStat.CurrentRec,
                inputStat.Delta,
                inputStat.NumPos,
                inputStat.TotalRecs,
                inputStat.CodePage,
                inputStat.TemplateLocale,
                inputStat.SortLocale);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1757
            bool isVerifyR1757 = (stat.SortType == inputStat.SortType)
                                && (stat.ContainerID == inputStat.ContainerID)
                                && (stat.CurrentRec == inputStat.CurrentRec)
                                && (stat.Delta == inputStat.Delta)
                                && (stat.NumPos == inputStat.NumPos)
                                && (stat.TotalRecs == inputStat.TotalRecs)
                                && (stat.CodePage == inputStat.CodePage)
                                && (stat.TemplateLocale == inputStat.TemplateLocale)
                                && (stat.SortLocale == inputStat.SortLocale);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1757,
                1757,
                @"[In NspiResortRestriction] If the server returns any return values other than ""Success"", the server MUST NOT modify the value of the parameter pStat.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiResortRestriction with the SortType field of the input parameter pStat containing any value other than SortTypeDisplayName or SortTypePhoneticDisplayName.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC21_ResortRestrictionWithSortTypeOtherThanDisplayNameOrPhoneticDisplayName()
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

            #region Call NspiDNToMId to map DN to Minimal Entry ID which will be used as the input parameter in the next step.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                CValues = 3,
                LppszA = new string[3]
                {
                    Common.GetConfigurationPropertyValue("User1Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User2Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User3Essdn", this.Site),
                }
            };

            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiDNToMId should return Success!");
            #endregion

            #region Call NspiResortRestriction with the SortType field of the input parameter pStat that is not SortTypeDisplayName or SortTypePhoneticDisplayName.

            uint reservedOfResortRestriction = 0;
            PropertyTagArray_r inmids = new PropertyTagArray_r();
            inmids = mids.Value;

            // The input parameter pStat is not SortTypeDisplayName or SortTypePhoneticDisplayName.
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName_RO;
            PropertyTagArray_r? outMIds = null;
            this.Result = this.ProtocolAdatper.NspiResortRestriction(reservedOfResortRestriction, ref stat, inmids, ref outMIds, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1185, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1185
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), (ErrorCodeValue)this.Result),
                1185,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the SortType field of the input parameter pStat contains any value other than ""SortTypeDisplayName"" or ""SortTypePhoneticDisplayName"", the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiCompareMIds returning InvalidBookmark.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC22_CompareMIdsFailedWithInvalidBookmark()
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

            #region Call NspiDNToMId to map DN to Minimal Entry ID which will be used as the input parameter in the next step.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                CValues = 2,
                LppszA = new string[2]
            };
            names.LppszA[0] = Common.GetConfigurationPropertyValue("User3Essdn", this.Site);
            names.LppszA[1] = Common.GetConfigurationPropertyValue("User1Essdn", this.Site);
            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiDNToMId should return Success!");
            #endregion

            #region Call NspiCompareMIds with the ContainerID field of the input parameter pStat set to an invalid value, so that the address book container specified by the invalid ContainerID field cannot be located.
            uint mid1 = mids.Value.AulPropTag[0];
            uint mid2 = mids.Value.AulPropTag[1];
            uint reservedOfCompareMIds = 0;
            int results;

            stat.ContainerID = (uint)MinimalEntryID.MID_CURRENT;
            this.Result = this.ProtocolAdatper.NspiCompareMIds(reservedOfCompareMIds, stat, mid1, mid2, out results);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1230");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1230
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidBookmark,
                this.Result,
                1230,
                @"[In NspiCompareMIds] [Server Processing Rules: Upon receiving message NspiCompareMIds, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the server is unable to locate the address book container specified by the ContainerID field in the input parameter pStat, the server MUST return the return value ""InvalidBookmark"".");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiCompareMIds returning GeneralFailure.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC23_CompareMIdsFailedWithGeneralFailure()
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

            #region Call NspiCompareMIds with invalid input parameters MId1 and MId2, so that the address book objects specified by invalid MId1 and MId2 cannot be located.
            uint reservedOfCompareMIds = 0;
            uint mid1 = 0x0; // A MID whose value is less than 0x10 is not used to specify any Address Book object.
            uint mid2 = 0x1;
            int results;

            this.Result = this.ProtocolAdatper.NspiCompareMIds(reservedOfCompareMIds, stat, mid1, mid2, out results, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1234");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1234
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.GeneralFailure,
                this.Result,
                1234,
                @"[In NspiCompareMIds] [Server Processing Rules: Upon receiving message NspiCompareMIds, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If the server is unable to locate the objects specified by the input parameters MId1 or MId2 in the table specified by the ContainerID field of the input parameter pStat, the server MUST return the return value ""GeneralFailure"".");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiModProps returning InvalidParameter.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC24_ModPropsFailedWithInvalidParameter()
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

            #region Call NspiGetSpecialTable method to get a special table from the server.
            uint version = 0;
            PropertyRowSet_r? rows;

            // Set flags to the value "NspiUnicodeStrings".
            uint flagsOfGetSpecialTable = (uint)NspiGetSpecialTableFlags.NspiUnicodeStrings;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable should return Success!");
            #endregion

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R750");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R750
            Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                version,
                750,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] If the client is requesting the rows of the server's address book hierarchy table and the server returns the value ""Success"", the server MUST set the output parameter lpVersion to the version of the server's address book hierarchy table.");

            #endregion

            #region Call NspiModProps with pPropTags set to NULL.
            uint reservedOfModProps = 0;
            PropertyRow_r rowOfModProps = rows.Value.ARow[0];
            PropertyTagArray_r? propTagsOfModProps = null;
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                this.Result = this.ProtocolAdatper.NspiModProps(reservedOfModProps, stat, propTagsOfModProps, rowOfModProps);

                #region Capture code
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1291: the value of the NspiModProps result is {0}", this.Result);

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1291
                Site.CaptureRequirementIfIsTrue(
                    propTagsOfModProps == null && this.Result == ErrorCodeValue.InvalidParameter,
                    1291,
                    @"[In NspiModProps] [Server Processing Rules: Upon receiving message NspiModProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If the input parameter pPropTags is NULL, the server MUST return the value ""InvalidParameter"".");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R903");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R903
                Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                    ErrorCodeValue.InvalidParameter,
                    this.Result,
                    "MS-OXCDATA",
                    903,
                    @"[In Error Codes] InvalidParameter(E_INVALIDARG, MAPI_E_INVALID_PARAMETER, ecInvalidParam, ecInvalidSession, ecBadBuffer,
                SYNC_E_INVALID_PARAMETER) will be returned, if an invalid parameter was passed to a remote procedure call (RPC).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R904");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R904
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    (uint)this.Result,
                    "MS-OXCDATA",
                    904,
                    @"[In Error Codes] The numeric value (hex) for error code InvalidParameter is 0x80070057, %x57.00.07.80.");
                #endregion
            }
            #endregion

            #region Call NspiGetSpecialTable method to get the special table again.
            PropertyRowSet_r? rows1;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable should return Success!");

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1284");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1284
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowSetEqual(rows, rows1),
                1284,
                @"[In NspiModProps] [Server Processing Rules: Upon receiving message NspiModProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the server returns any return value other than ""Success"", the server MUST NOT modify any properties of any objects in the address book.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R750001");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R750001
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowSetEqual(rows, rows1),
                750001,
                @"[In NspiGetSpecialTable] The Exchange server behavior is considered special as the Ipversion here does not impact any search results. ");
            #endregion
            #endregion

            #region Call NspiModProps again with the CurrentRec field of the input parameter pStat set to an invalid value, so that the server cannot locate the object specified by the invalid CurrentRec.
            reservedOfModProps = 0;
            BinaryArray_r certificate = new BinaryArray_r();
            rowOfModProps = new PropertyRow_r
            {
                LpProps = new PropertyValue_r[1]
            };
            rowOfModProps.LpProps[0].PropTag = (uint)AulProp.PidTagAddressBookX509Certificate;
            rowOfModProps.LpProps[0].Value.MVbin = certificate;

            PropertyTagArray_r instanceOfModProps = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagAddressBookX509Certificate
                }
            };
            propTagsOfModProps = instanceOfModProps;

            // Set CurrentRec to a value which server is unable to locate.
            stat.CurrentRec = (uint)MinimalEntryID.MID_CURRENT;
            this.Result = this.ProtocolAdatper.NspiModProps(reservedOfModProps, stat, propTagsOfModProps, rowOfModProps);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1293");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1293
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidParameter,
                this.Result,
                1293,
                @"[In NspiModProps] [Server Processing Rules: Upon receiving message NspiModProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] If the server is unable to locate the object specified by the CurrentRec field of the input parameter pStat, the server MUST return the value ""InvalidParameter"".");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiResolveNames returning InvalidBookmark.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC25_ResolveNamesFailedWithInvalidBookmark()
        {
            this.CheckProductSupported();
            if (this.Transport == "mapi_http")
            {
                Site.Assert.Inconclusive("This case cannot run, since MAPIHTTP transport does not support ResolveNames operation for 8-bit character set string.");
            }

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

            #region Call NspiResolveNames with the ContainerID field of the input parameter pStat set to an invalid value.
            uint reservedOfResolveNames = 1;
            StringsArray_r strArray;
            strArray.CValues = 1;
            strArray.LppszA = new string[strArray.CValues];
            strArray.LppszA[0] = Common.GetConfigurationPropertyValue("User1Name", this.Site);

            PropertyTagArray_r? propTags = null;
            PropertyTagArray_r? mids;
            PropertyRowSet_r? rowOfResolveNames;

            stat.ContainerID = (uint)MinimalEntryID.MID_CURRENT;
            this.Result = this.ProtocolAdatper.NspiResolveNames(reservedOfResolveNames, stat, propTags, strArray, out mids, out rowOfResolveNames);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1367");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1367
            Site.CaptureRequirementIfIsNull(
                mids,
                1367,
                @"[In NspiResolveNames] [Server Processing Rules: Upon receiving message NspiResolveNames, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the server returns any return value other than ""Success"", the server MUST return the value NULL in the return parameters ppMIds.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1372");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1372
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidBookmark,
                this.Result,
                1372,
                @"[In NspiResolveNames] [Server Processing Rules: Upon receiving message NspiResolveNames, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If the server is unable to locate the address book container specified by the ContainerID field in the input parameter pStat, the server MUST return the return value ""InvalidBookmark"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1760");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1760
            // When test code reaches here, server returns a value that is not "Success". So only whether rowOfResolveNames is null or not needs to be determined.
            Site.CaptureRequirementIfIsNull(
                rowOfResolveNames,
                1760,
                @"[In NspiResolveNames] If the server returns any return value other than ""Success"", the server MUST return the value NULL
                in the return parameters ppRows.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiResolveNames with the input parameter Reserved that contains any value other than 0.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC26_ResolveNamesWithReservedNonZero()
        {
            this.CheckProductSupported();
            if (this.Transport == "mapi_http")
            {
                Site.Assert.Inconclusive("This case cannot run, since MAPIHTTP transport does not support ResolveNames operation for 8-bit character set string.");
            }

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
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return success!");
            #endregion

            #region Call NspiResolveNames with the input parameter Reserved set to 1.
            uint reservedOfResolveNames = 0x1;
            StringsArray_r strArray;
            strArray.CValues = 2;
            strArray.LppszA = new string[strArray.CValues];
            strArray.LppszA[0] = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            strArray.LppszA[1] = string.Empty;

            PropertyTagArray_r? propTags = null;
            PropertyTagArray_r? mids;
            PropertyRowSet_r? rowOfResolveNames;

            this.Result = this.ProtocolAdatper.NspiResolveNames(reservedOfResolveNames, stat, propTags, strArray, out mids, out rowOfResolveNames, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1363, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1363
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1363,
                @"[In NspiResolveNames] [Server Processing Rules: Upon receiving message NspiResolveNames, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the input parameter Reserved contains any value other than 0, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiResolveNamesW operation returning InvalidBookmark.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC27_ResolveNamesWFailedWithInvalidBookmark()
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
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return success!");
            #endregion

            #region Call NspiResolveNamesW with the ContainerID field of the input parameter pStat set to an invalid value, so that the address book container specified by the invalid ContainerID field cannot be located.
            uint reservedOfResolveNamesW = 0;
            WStringsArray_r wstrArray;
            wstrArray.CValues = 2;
            wstrArray.LppszW = new string[wstrArray.CValues];
            wstrArray.LppszW[0] = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            wstrArray.LppszW[1] = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            PropertyTagArray_r? propTags = null;
            PropertyTagArray_r? mids;
            PropertyRowSet_r? rowOfResolveNamesW;

            stat.ContainerID = (uint)MinimalEntryID.MID_CURRENT;
            this.Result = this.ProtocolAdatper.NspiResolveNamesW(reservedOfResolveNamesW, stat, propTags, wstrArray, out mids, out rowOfResolveNamesW);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1418");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1418
            Site.CaptureRequirementIfIsNull(
                mids,
                1418,
                @"[In NspiResolveNamesW] [Server Processing Rules: Upon receiving message NspiResolveNamesW, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the server returns any return value other than ""Success"", the server MUST return the value NULL in the return parameters ppMIds.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1761");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1761
            Site.CaptureRequirementIfIsNull(
                rowOfResolveNamesW,
                1761,
                @"[In NspiResolveNamesW] If the server returns any return value other than ""Success"", the server MUST return the value NULL
                in the return parameters ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1423");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1423
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidBookmark,
                this.Result,
                1423,
                @"[In NspiResolveNamesW] [Server Processing Rules: Upon receiving message NspiResolveNamesW, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If the server is unable to locate the address book container specified by the ContainerID field in the input parameter pStat, the server MUST return the return value ""InvalidBookmark"".");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiResolveNamesW with the input parameter Reserved that contains any value other than 0.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC28_ResolveNamesWWithReservedNonZero()
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
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return success!");
            #endregion

            #region Call NspiResolveNamesW with the input parameter Reserved set to 1.
            uint reservedOfResolveNamesW = 0x1;

            WStringsArray_r wstrArray;
            wstrArray.CValues = 2;
            wstrArray.LppszW = new string[wstrArray.CValues];
            wstrArray.LppszW[0] = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            wstrArray.LppszW[1] = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            PropertyTagArray_r? propTags = null;
            PropertyTagArray_r? mids;
            PropertyRowSet_r? rowOfResolveNamesW;

            this.Result = this.ProtocolAdatper.NspiResolveNamesW(reservedOfResolveNamesW, stat, propTags, wstrArray, out mids, out rowOfResolveNamesW, false);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1414, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1414
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1414,
                @"[In NspiResolveNamesW] [Server Processing Rules: Upon receiving message NspiResolveNamesW, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the input parameter Reserved contains any value other than 0, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetTemplateInfo returning InvalidCodepage.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC29_GetTemplateInfoFailedWithInvalidCodePage()
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
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return success!");
            #endregion

            #region Call NspiGetTemplateInfo with the CodePage field in the dwCodePage input parameter set to CP_WINUNICODE.
            uint flagsOfGetTemplateInfo = (uint)NspiGetTemplateInfoFlag.TI_HELPFILE_NAME;
            uint type = 0;
            string dn = null;
            uint codePage = (uint)RequiredCodePage.CP_WINUNICODE;
            uint locateID = stat.TemplateLocale;
            PropertyRow_r? data;

            this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data);

            #region Capture

            this.VerifyWhetherppDataIsNullForNspiGetTemplateInfo(data);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1467");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1467
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidCodepage,
                this.Result,
                1467,
                @"[In NspiGetTemplateInfo] [Server Processing Rules: Upon receiving message NspiGetTemplateInfo, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the codepage specified in the dwCodePage input parameter has the value CP_WINUNICODE, the server MUST return the value ""InvalidCodePage"".");

            #endregion Capture
            #endregion

            #region Call NspiGetTemplateInfo with the CodePage field in the dwCodePage input parameter set to an invalid value that the server cannot recognize.
            // If the server does not recognize the CodePage specified in the dwCodePage input parameter as a supported code page,
            // the server MUST return the value "InvalidCodePage".
            codePage = uint.Parse(Constants.UnrecognizedCodePage);
            this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data);

            #region Capture

            this.VerifyWhetherppDataIsNullForNspiGetTemplateInfo(data);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1469");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1469
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidCodepage,
                this.Result,
                1469,
                @"[In NspiGetTemplateInfo] [Server Processing Rules: Upon receiving message NspiGetTemplateInfo, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] If the server does not recognize the codepage specified in the dwCodePage input parameter as a supported code page, the server MUST return the value ""InvalidCodePage"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R973");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R973
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidCodepage,
                this.Result,
                "MS-OXCDATA",
                973,
                @"[In Error Codes] InvalidCodepage(MAPI_E_UNKNOWN_CPID) will be returned, if the server is not configured to support the code
                page requested by the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R974");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R974
            Site.CaptureRequirementIfAreEqual<uint>(
                 0x8004011E,
                (uint)this.Result,
                "MS-OXCDATA",
                974,
                @"[In Error Codes] The numeric value (hex) for error code InvalidCodepage is 0x8004011E, %x1E.01.04.80.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetTemplateInfo returning InvalidLocale.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC30_GetTemplateInfoFailedWithInvalidLocale()
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
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return success!");
            #endregion

            #region Call NspiGetTemplateInfo with the input parameter pDN set to NULL and the input parameter dwLocaleID set to an invalid value, so that the server cannot locate the object specified by the invalid dwLocaleID.
            string dn = null;
            uint locateID = 0x0; // An invalid value which does not specify any LCID according to [MS-LCID].
            uint flagsOfGetTemplateInfo = (uint)NspiGetTemplateInfoFlag.TI_HELPFILE_NAME;
            uint type = (uint)DisplayTypeValue.DT_MAILUSER;
            uint codePage = stat.CodePage;
            PropertyRow_r? data;

            this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data);

            #region Capture

            this.VerifyWhetherppDataIsNullForNspiGetTemplateInfo(data);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1476");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1476
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidLocale,
                this.Result,
                1476,
                @"[In NspiGetTemplateInfo] [Server Processing Rules: Upon receiving message NspiGetTemplateInfo, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If the server is unable to locate a specific object based on these constraints [constraints 1-4], the server MUST return the value ""InvalidLocale"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R975");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R975
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.InvalidLocale,
                this.Result,
                "MS-OXCDATA",
                975,
                @"[In Error Codes] InvalidLocale(MAPI_E_UNKNOWN_LCID) will be returned, if the server is not configured to support the locale
                requested by the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R976");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R976
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004011F,
                (uint)this.Result,
                "MS-OXCDATA",
                976,
                @"[In Error Codes] The numeric value (hex) for error code InvalidLocale is 0x8004011F, %x1F.01.04.80.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetProps returning failure response.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC31_NspiGetPropsFailure()
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

            this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return success!");
            #endregion

            #region Call NspiGetProps method with invalid ContainerID.
            PropertyTagArray_r prop = new PropertyTagArray_r
            {
                CValues = 1
            };
            prop.AulPropTag = new uint[prop.CValues];
            prop.AulPropTag[0] = (uint)AulProp.PidTagUserX509Certificate;
            PropertyTagArray_r? propTags = prop;

            uint flagsOfGetProps = (uint)RetrievePropertyFlag.fEphID;
            PropertyRow_r? rows;

            stat.ContainerID = 1000;  // Invalid containerID
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);

            bool isR2003005Verified = false;
            if (Common.IsRequirementEnabled(2003005, this.Site))
            {
                if (ErrorCodeValue.ErrorsReturned == this.Result)
                {
                    isR2003005Verified = true;
                }
            }
            else
            {
                if (ErrorCodeValue.InvalidBookmark == this.Result)
                {
                    isR2003005Verified = true;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R2003005");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R2003005
            Site.CaptureRequirementIfIsTrue(
                isR2003005Verified,
                2003005,
                @"[In Appendix A: Product Behavior] Implementation does return the value ""ErrorsReturned"" (0x00040380). <4> Section 3.1.4.1.7:  Exchange 2010 SP3, Exchange 2013, and Exchange 2016 return ""ErrorsReturned"" (0x00040380).");
            #endregion

            #region Call NspiUpdateStat to update the STAT block to make CurrentRec point to the first row of the table.
            uint reserved = 0;
            int? delta = 1;
            stat.ContainerID = 0; // Reset the valid containerID
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return success!");
            #endregion

            #region Call NspiGetProps method with dwFlags set to fEphID.
            
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.IsNotNull(rows, "rows should not be null. The row number is {0}.", rows == null ? 0 : rows.Value.CValues);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R910");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R910
            Site.CaptureRequirementIfAreNotEqual<int>(
                0,
                rows.Value.LpProps[0].Value.Err,
                910,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 12] [If a property in the proptag list has no value on the object specified by the CurrentRec field] The server MUST set the aulPropTag member corresponding to the proptag with no value with the proptag that has no value with the PtypErrorCode property type.");
            #endregion

            #region Call NspiQueryRows with propTags contains PidTagContainerContents and PidTagContainerFlags.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            uint tableCount = 0;
            uint[] table = null;
            uint count = 10;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[]
                {
                    (uint)AulProp.PidTagUserX509Certificate
                }
            };
            propTags = propTagsInstance;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);

            foreach (PropertyRow_r propertyRow in rowsOfQueryRows.Value.ARow)
            {
                Site.Assert.AreEqual<uint>(
                    0x0000000A,
                    propertyRow.LpProps[0].PropTag & 0x0000000A,
                    "The property type fields of the property should set to 0x0000000A (PtypErrorCode)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R986");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R986
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                986,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 15] If the server has no rows that satisfy this query [the query specified by method NspiQueryRows], the server MUST return the value ""Success"".");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R985
            // If NspiQueryRows returns Success and the property type fields of the property are all set to 0x0000000A (PtypErrorCode), MS-OXNSPI_R985 can be verified directly.
            Site.CaptureRequirement(
                985,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 15] If the server has no rows that satisfy this query, the server MUST return the value ""Success"" and place a PropertyRowSet_r with rows according to the input parameter ""Count"" in the output parameter ppRows, in which the property type fields of the property are all set to 0x0000000A (PtypErrorCode).");
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NotFound error code.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC32_VerifyNotFoundErrorCode()
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
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return success!");

            #endregion

            #region Call NspiGetMatches to get valid Minimal Entry IDs and rows.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            Restriction_r res_r = new Restriction_r
            {
                Rt = 0x8,
                Res =
                    new RestrictionUnion_r
                    {
                        ResExist =
                            new ExistRestriction_r
                            {
                                Reserved1 = 0,
                                Reserved2 = 0,
                                PropTag = (uint)AulProp.PidTagDisplayName
                            }
                    }
            };
            Restriction_r? filter = res_r;

            PropertyName_r? propNameOfGetMatches = null;
            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagDisplayName,
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return success!");
            #endregion

            #region Call NspiGetProps to get the property PidTagUserX509Certificate which doesn't not exist in the address book object.
            propTags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagUserX509Certificate,
                }
            };
            propTagsOfGetMatches = propTags;
            PropertyRow_r? rows;

            this.Result = this.ProtocolAdatper.NspiGetProps(flags, stat, propTagsOfGetMatches, out rows);
            uint errorcode = (uint)rows.Value.LpProps[0].Value.Err;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2055");

            // Verify MS-OXNSPI requirement: MS-OXCDATA_R2055
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010f,
                errorcode,
                "MS-OXCDATA",
                2055,
                @"[In Additional Error Codes] The numeric value (hex) for error code NotFound is 0x8004010F, %x0F.01.04.80.");

            // Verify MS-OXNSPI requirement: MS-OXCDATA_R2054
            // If the error code %x0F.01.04.80 has been captured in MS-OXCDATA_R2055, MS-OXCDATA_R2054 can be verified directly.
            Site.CaptureRequirement(
                "MS-OXCDATA",
                2054,
                @"[In Additional Error Codes] NotFound (MAPI_E_NOT_FOUND) will be returned, On get, indicates that the property or column has no value for this object.");
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiModProps returning failure response.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC33_NspiModPropsFaliure()
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
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return success!");

            #endregion

            #region Call NspiGetMatches to get valid Minimal Entry IDs and rows.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            // Create Restriction_r structure to use the display name of specific user as the filter parameter of NspiGetMatches method.
            Restriction_r propertyRestriction = new Restriction_r
            {
                Rt = 0x04,
                Res = new RestrictionUnion_r
                {
                    ResProperty = new Propertyrestriction_r
                    {
                        Relop = 0x04
                    }
                }
            };
            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = 0
            };
            string userName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(userName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(userName + "\0");
            }

            propertyRestriction.Res.ResProperty.Prop = new PropertyValue_r[] { target };
            propertyRestriction.Res.ResProperty.PropTag = (uint)AulProp.PidTagDisplayName;

            Restriction_r? filter = propertyRestriction;

            PropertyName_r? propNameOfGetMatches = null;
            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagDisplayName,
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            #endregion

            #region Call NspiModProps method with specific PidTagAddressBookX509Certificate property value.

            // Get user name.         
            uint reservedOfModProps = 0xff;
            PropertyRow_r rowOfModProps = new PropertyRow_r
            {
                LpProps = new PropertyValue_r[1]
            };
            rowOfModProps.LpProps[0].PropTag = (uint)AulProp.PidTagDisplayName;
            rowOfModProps.LpProps[0].Value.LpszA = System.Text.Encoding.UTF8.GetBytes(userName.ToUpper(System.Globalization.CultureInfo.CurrentCulture));

            PropertyTagArray_r instanceOfModProps = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagDisplayName
                }
            };
            PropertyTagArray_r? propTagsOfModProps = instanceOfModProps;

            // Set the CurrentRec field with the minimal entry ID of mail user name.
            for (int i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                string name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[0].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(userName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    stat.CurrentRec = outMIds.Value.AulPropTag[i];
                    break;
                }
            }

            this.Result = this.ProtocolAdatper.NspiModProps(reservedOfModProps, stat, propTagsOfModProps, rowOfModProps);

            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiSeekEntries returning failure response.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC34_NspiSeekEntriesFaliure()
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
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiBind should return success!");
            #endregion

            #region Call NspiGetMatches method with specified filter and prop tags.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            Restriction_r contentRestriction = new Restriction_r
            {
                Rt = 0x03,
                Res = new RestrictionUnion_r
                {
                    ResContent = new ContentRestriction_r
                    {
                        FuzzyLevel = 0x00010002
                    }
                }
            };

            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = 0
            };
            string displayName = Common.GetConfigurationPropertyValue("DistributionListName", this.Site);
            target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName + "\0");
            }

            contentRestriction.Res.ResContent.Prop = new PropertyValue_r[] { target };
            contentRestriction.Res.ResContent.PropTag = (uint)AulProp.PidTagDisplayName;
            Restriction_r? filter = contentRestriction;

            PropertyTagArray_r propTags1 = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[]
                {
                    (uint)AulProp.PidTagDisplayName
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags1;

            PropertyName_r? propNameOfGetMatches = null;
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return success!");
            #endregion

            #region Call NspiSeekentries method with input parameter lpETable is not set to null.
            PropertyValue_r propertyTarget = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = (uint)0x00
            };
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                propertyTarget.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);
            }
            else
            {
                propertyTarget.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName + "\0");
            }

            target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);
            propertyTarget.PropTag = (uint)AulProp.PidTagDisplayName;
            PropertyTagArray_r? table = outMIds;
            PropertyTagArray_r? propertyTags = null;
            PropertyRowSet_r? rowsOfSeekEntries;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reserved1, ref stat, propertyTarget, table, propertyTags, out rowsOfSeekEntries, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1020, the value of the result is {0}", this.Result);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1020
            Site.CaptureRequirementIfIsTrue(
                Enum.IsDefined(typeof(ErrorCodeValue), this.Result),
                1020,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the input parameter lpETable is not NULL and does not contain an Explicit Table both containing a restriction of the table specified by the input parameter pStat and sorted as specified by the SortType field of the input parameter pStat, the server MUST return one of the return values [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] specified in section 2.2.1.2.");
            #endregion

            #region Call NspiSeekentries method and the input parameter STAT block specifies SortTypePhoneticDisplayName and the property specified in the input parameter pTarget is PidTagEntryId.
            propTags1 = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[]
                {
                    (uint)AulProp.PidTagEntryId
                }
            };
            propertyTags = propTags1;

            stat.SortType = (uint)TableSortOrder.SortTypePhoneticDisplayName;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reserved1, ref stat, propertyTarget, table, propertyTags, out rowsOfSeekEntries, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1045");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1045
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.GeneralFailure,
                this.Result,
                1045,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 12] If the SortType field in the input parameter pStat is SortTypePhoneticDisplayName and the property specified in the input parameter pTarget is anything other than PidTagAddressBookPhoneticDisplayName (with either the Property Type PtypString8 or PtypString), the server MUST return the value GeneralFailure.");
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetMatches returning GeneralFailure.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S05_TC35_GetMatchesFailedWithGeneralFailure()
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

            #region Call NspiGetMatches with the Filter field set to null and the STAT block specifies an invalid position.
            uint reserved1 = 0;
            uint reserver2 = 0;
            uint requested = Constants.GetMatchesRequestedRowNumber;
            PropertyName_r? propNameOfGetMatches = null;
            PropertyTagArray_r? proReserved = null;
            Restriction_r? filter = null;
            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagEntryId,
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches, false);

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1138");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1138
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.GeneralFailure,
                this.Result,
                1138,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 10] If the server is unable to locate the object, the server MUST return the value ""GeneralFailure"".");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        #region Private methods

        /// <summary>
        /// Verify parameters related to method NspiSeekEntries.
        /// </summary>
        /// <param name="result">A DWORD value that specifies the return status of the method.</param>
        /// <param name="rowsOfSeekEntries">A PropertyRowSet_r value which contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="inputStat">The STAT parameter before calling NspiSeekEntries.</param>
        /// <param name="stat">The STAT parameter after calling NspiSeekEntries.</param>
        private void VerifyParametersRelatedWithNspiSeekEntries(ErrorCodeValue result, PropertyRowSet_r? rowsOfSeekEntries, STAT inputStat, STAT stat)
        {
            if (result != ErrorCodeValue.Success)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1630");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1630
                Site.CaptureRequirementIfIsNull(
                    rowsOfSeekEntries,
                    1630,
                    @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] If the server returns any return values other than ""Success"", the server MUST return a NULL for the output parameter ppRows");

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXNSPI_R1631: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
                    "the SortType of inputStat is {9}, the ContainerID of inputStat is {10}, the CurrentRec of inputStat is {11}, the Delta of inputStat is {12}, the NumPos of inputStat is {13}, the TotalRecs of inputStat is {13}, the CodePage of inputStat is {14}, the TemplateLocale of inputStat is {15}, the SortLocale of inputStat is {16}",
                    stat.SortType,
                    stat.ContainerID,
                    stat.CurrentRec,
                    stat.Delta,
                    stat.NumPos,
                    stat.TotalRecs,
                    stat.CodePage,
                    stat.TemplateLocale,
                    stat.SortLocale,
                    inputStat.SortType,
                    inputStat.ContainerID,
                    inputStat.CurrentRec,
                    inputStat.Delta,
                    inputStat.NumPos,
                    inputStat.TotalRecs,
                    inputStat.CodePage,
                    inputStat.TemplateLocale,
                    inputStat.SortLocale);

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1631
                bool isVerifyR1631 = (stat.SortType == inputStat.SortType)
                                    && (stat.ContainerID == inputStat.ContainerID)
                                    && (stat.CurrentRec == inputStat.CurrentRec)
                                    && (stat.Delta == inputStat.Delta)
                                    && (stat.NumPos == inputStat.NumPos)
                                    && (stat.TotalRecs == inputStat.TotalRecs)
                                    && (stat.CodePage == inputStat.CodePage)
                                    && (stat.TemplateLocale == inputStat.TemplateLocale)
                                    && (stat.SortLocale == inputStat.SortLocale);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1631,
                    1631,
                    @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] [If the server returns any return values other than ""Success"", the server] MUST NOT modify the value of the parameter pStat.");
            }
        }

        /// <summary>
        /// Verify parameters related to method NspiGetMatches.
        /// </summary>
        /// <param name="result">A DWORD value that specifies the return status of the method.</param>
        /// <param name="outMIds">A PropertyTagArray_r value which holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="rowsOfGetMatches">A PropertyRowSet_r value which contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="stat">A reference to a STAT block describing a logical position in a specific address book container.</param>
        /// <param name="inputStat">The input parameter of STAT.</param>
        private void VerifyParametersRelatedWithNspiGetMatches(ErrorCodeValue result, PropertyTagArray_r? outMIds, PropertyRowSet_r? rowsOfGetMatches, STAT stat, STAT inputStat)
        {
            if (result != ErrorCodeValue.Success)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1635");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1635
                Site.CaptureRequirementIfIsNull(
                    outMIds,
                    1635,
                    @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] If the server returns any return values other than ""Success"", the server MUST return a NULL for the output parameters ppOutMIds.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1756");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1756
                Site.CaptureRequirementIfIsNull(
                    rowsOfGetMatches,
                    1756,
                    @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] If the server returns any return values other than ""Success"", the server MUST return a NULL for the output parameters ppRows.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1636");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1636
                bool isVerifyR1636 = (stat.SortType == inputStat.SortType)
                                    && (stat.ContainerID == inputStat.ContainerID)
                                    && (stat.CurrentRec == inputStat.CurrentRec)
                                    && (stat.Delta == inputStat.Delta)
                                    && (stat.NumPos == inputStat.NumPos)
                                    && (stat.TotalRecs == inputStat.TotalRecs)
                                    && (stat.CodePage == inputStat.CodePage)
                                    && (stat.TemplateLocale == inputStat.TemplateLocale)
                                    && (stat.SortLocale == inputStat.SortLocale);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1636,
                    1636,
                    @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] [If the server returns any return values other than ""Success"",] and [the server] MUST NOT modify the value of the parameter pStat.");
            }
        }

        /// <summary>
        /// Check whether ppData field is null.
        /// </summary>
        /// <param name="data">Contain the information requested by calling NspiGetTemplateInfo.</param>
        private void VerifyWhetherppDataIsNullForNspiGetTemplateInfo(PropertyRow_r? data)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1462");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1462
            Site.CaptureRequirementIfIsNull(
                data,
                1462,
                @"[In NspiGetTemplateInfo] [Server Processing Rules: Upon receiving message NspiGetTemplateInfo, the server MUST process the data from the message subject to the following constraints:] [Constraint 1] If the server returns any return value other than ""Success"", the server MUST return the value NULL in the return parameters ppData.");
        }

        /// <summary>
        /// Check whether STAT is modified.
        /// </summary>
        /// <param name="statSave">The STAT parameter before calling NspiUpdateStat.</param>
        /// <param name="stat">The STAT parameter after calling NspiUpdateStat.</param>
        private void VerifyWhetherpStatIsModifiedForNspiUpdateStat(STAT statSave, STAT stat)
        {
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R784: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
                "the SortType of statSave is {9}, the ContainerID of statSave is {10}, the CurrentRec of statSave is {11}, the Delta of statSave is {12}, the NumPos of statSave is {13}, the TotalRecs of statSave is {13}, the CodePage of statSave is {14}, the TemplateLocale of statSave is {15}, the SortLocale of statSave is {16}",
                stat.SortType,
                stat.ContainerID,
                stat.CurrentRec,
                stat.Delta,
                stat.NumPos,
                stat.TotalRecs,
                stat.CodePage,
                stat.TemplateLocale,
                stat.SortLocale,
                statSave.SortType,
                statSave.ContainerID,
                statSave.CurrentRec,
                statSave.Delta,
                statSave.NumPos,
                statSave.TotalRecs,
                statSave.CodePage,
                statSave.TemplateLocale,
                statSave.SortLocale);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R784
            bool isVerifyR784 = (stat.CodePage == statSave.CodePage) && (stat.ContainerID == statSave.ContainerID)
                && (stat.CurrentRec == statSave.CurrentRec) && (stat.Delta == statSave.Delta)
                && (stat.NumPos == statSave.NumPos) && (stat.SortLocale == statSave.SortLocale)
                && (stat.SortType == statSave.SortType) && (stat.TemplateLocale == statSave.TemplateLocale);

            // The Open Specification defines many error codes other than "Success", but in this operation it only describes when the server will return "InvalidBookmark" and "Not Found".
            // If isVerifyR784 is true and STAT is equal to statSave, the server does not modify STAT.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR784,
                784,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 2] If the server returns any return value other than ""Success"", the server MUST NOT modify the output parameter pStat.");
        }
        #endregion
    }
}