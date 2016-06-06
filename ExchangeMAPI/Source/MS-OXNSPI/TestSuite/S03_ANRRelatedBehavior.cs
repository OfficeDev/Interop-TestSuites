namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class contains all the test cases designed to test the server behavior for the NSPI calls related to Ambiguous Name Resolution process.
    /// </summary>
    [TestClass]
    public class S03_ANRRelatedBehavior : TestSuiteBase
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
        ///  This test case is designed to verify the requirements related to NspiResolveNames operation returning success.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S03_TC01_ResolveNamesSuccess()
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

            #region Call NspiResolveNames method with a specific list of names and propertyTag value with null.
            StringsArray_r strArray;
            strArray.CValues = 5;
            strArray.LppszA = new string[strArray.CValues];
            strArray.LppszA[0] = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            strArray.LppszA[1] = string.Empty;
            strArray.LppszA[2] = Common.GetConfigurationPropertyValue("AmbiguousName", this.Site);
            strArray.LppszA[3] = null;
            strArray.LppszA[4] = Constants.UnresolvedName;

            PropertyTagArray_r? propTags = null;
            PropertyTagArray_r? mids;
            PropertyRowSet_r? rowOfResolveNames;

            this.Result = this.ProtocolAdatper.NspiResolveNames((uint)0, stat, propTags, strArray, out mids, out rowOfResolveNames);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1392");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1392
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1392,
                @"[In NspiResolveNames] [Server Processing Rules: Upon receiving message NspiResolveNames, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] If no other return values have been specified by these constraints [constraints 1-8], the server MUST return the return value ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1342");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1342
            // If the mids which contains Minimal Entry ID returned from the server is not null, 
            // it illustrates that the server must report the Minimal Entry ID that is the result of the ANR process.
            this.Site.CaptureRequirementIfIsNotNull(
                mids,
                1342,
                @"[In NspiResolveNames] The server reports the Minimal Entry ID that is the result of the ANR process.");

            this.VerifyPropertyRowSetIsNotNullForNspiResolveNames(rowOfResolveNames);

            this.VerifyIsRESOLVEDMIDInANRMatchString(mids.Value.AulPropTag[0]);
            this.VerifyIsUNRESOLVEDMIDInANREmptyString(mids.Value.AulPropTag[1]);
            this.VerifyIsAMBIGUOUSMIDInANRAmbiguousString(mids.Value.AulPropTag[2]);
            this.VerifyIsUNRESOLVEDMIDInANRNullString(mids.Value.AulPropTag[3]);
            this.VerifyIsUNRESOLVEDMIDInANRNotFound(mids.Value.AulPropTag[4]);
            this.VerifyIsResultOfANRProcessIsMID(mids);

            // Since if the whole ANR results are verified and the NspiResolveNames is invoked, 
            // it must take a set of string values in an 8-bit character set and perform ANR on those strings, 
            // this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1341,
                @"[In NspiResolveNames] The NspiResolveNames method takes a set of string values in an 8-bit character set and performs ANR (as specified in section 3.1.4.7) on those strings.");

            // Since all the returned Minimal Entry IDs are verified according to the order of the input string array, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1375,
                @"[In NspiResolveNames] [Server Processing Rules: Upon receiving message NspiResolveNames, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] These Minimal Entry IDs [IDs which server constructs and returns to client] are those that result from applying the ANR process, as specified in section 3.1.4.7, to the strings in the input parameter paStr.");

            #endregion Capture
            #endregion

            #region Call NspiResolveNames method with a specific list of names and propertyTag value without null.
            PropertyTagArray_r propTagsWithProperties = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagObjectType
                }
            };

            this.Result = this.ProtocolAdatper.NspiResolveNames((uint)0, stat, propTagsWithProperties, strArray, out mids, out rowOfResolveNames);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiResolveNames should return Success!");

            PropertyRow_r rowValue = rowOfResolveNames.Value.ARow[0];

            string resolvedName = System.Text.UnicodeEncoding.UTF8.GetString(rowValue.LpProps[1].Value.LpszA);

            #region Capture

            // Check whether MID_RESOLVED is returned.
            bool isMIDRESOLVEDContained = false;
            for (int i = 0; i < mids.Value.CValues; i++)
            {
                if (mids.Value.AulPropTag[i] == 0x0000002)
                {
                    isMIDRESOLVEDContained = true;
                    break;
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1343");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1343
            // The rowOfResolveNames which contains a PropertyRowSet_r structure is returned from the server. If it is not null and MID_RESOLVED is returned,
            // it illustrates that certain property values are returned for any valid Minimal Entry IDs identified by the ANR process.
            this.Site.CaptureRequirementIfIsTrue(
                rowOfResolveNames != null && isMIDRESOLVEDContained,
                1343,
                @"[In NspiResolveNames] Certain property values are returned for any valid Minimal Entry IDs identified by the ANR process.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1349, the display name of the resolved address book for paStr {0} is {1}.",
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                resolvedName);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1349
            this.Site.CaptureRequirementIfIsTrue(
                resolvedName.Equals(Common.GetConfigurationPropertyValue("User1Name", this.Site), StringComparison.CurrentCultureIgnoreCase),
                1349,
                "[In NspiResolveNames] [paStr] Specifies the values the client is requesting the server to do ANR on.");

            Site.Assert.AreEqual<uint>((uint)AulProp.PidTagEntryId, rowOfResolveNames.Value.ARow[0].LpProps[0].PropTag, "The first property returned should be PidTagEntryId. Now the returned property is {0}.", rowOfResolveNames.Value.ARow[0].LpProps[0].PropTag);
            Site.Assert.AreEqual<uint>((uint)AulProp.PidTagDisplayName, rowOfResolveNames.Value.ARow[0].LpProps[1].PropTag, "The second property returned should be PidTagDisplayName. Now the returned property is {0}.", rowOfResolveNames.Value.ARow[0].LpProps[1].PropTag);
            Site.Assert.AreEqual<uint>((uint)AulProp.PidTagObjectType, rowOfResolveNames.Value.ARow[0].LpProps[2].PropTag, "The third property returned should be PidTagObjectType. Now the returned property is {0}.", rowOfResolveNames.Value.ARow[0].LpProps[2].PropTag);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1347");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1347
            // Since the null value of pPropTags has been set in step "Call NspiResolveNames method with a specific list of names", 
            // and the above three asserts ensure that it is a reference to a PropertyTagArray_r value containing a list of the proptags of the columns that the client requests to be returned for each row.
            // So R1347 can be captured directly.
            this.Site.CaptureRequirement(
                1347,
                @"[In NspiResolveNames] pPropTags: The value NULL or a reference to a PropertyTagArray_r value containing a list of the proptags of the columns that the client requests to be returned for each row returned.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        ///  This test case is designed to verify the requirements related to NspiResolveNamesW operation returning success.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S03_TC02_ResolveNamesWSuccess()
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

            #region Call NspiResolveNamesW method with a specific list of names and propertyTag value with null.
            uint reservedOfResolveNamesW = 0;
            WStringsArray_r wstrArray;
            wstrArray.CValues = 5;
            wstrArray.LppszW = new string[wstrArray.CValues];
            wstrArray.LppszW[0] = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            wstrArray.LppszW[1] = string.Empty;
            wstrArray.LppszW[2] = Common.GetConfigurationPropertyValue("AmbiguousName", this.Site);
            wstrArray.LppszW[3] = null;
            wstrArray.LppszW[4] = Constants.UnresolvedName;

            PropertyTagArray_r? propTags = null;
            PropertyTagArray_r? mids;
            PropertyRowSet_r? rowOfResolveNamesW;

            this.Result = this.ProtocolAdatper.NspiResolveNamesW(reservedOfResolveNamesW, stat, propTags, wstrArray, out mids, out rowOfResolveNamesW);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1443");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1443
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1443,
                @"[In NspiResolveNamesW] [Server Processing Rules: Upon receiving message NspiResolveNamesW, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] If no other return values have been specified by these constraints [constraints 1-8], the server MUST return the return value ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1394");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1394
            // If the mids which contains Minimal Entry ID returned from the server is not null, 
            // it illustrates that the server must report the Minimal Entry ID that is the result of the ANR process.
            this.Site.CaptureRequirementIfIsNotNull(
                mids,
                1394,
                @"[In NspiResolveNamesW] The server reports the Minimal Entry IDs that are the result of the ANR process.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1650");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1650
            Site.CaptureRequirementIfIsNotNull(
                mids,
                1650,
                "[In NspiResolveNamesW] [ppMIds] On return, contains a list of Minimal Entry IDs that match the array of strings, as specified in the input parameter paWStr.");

            this.VerifyPropertyRowSetIsNotNullForNspiResolveNamesW(rowOfResolveNamesW);

            this.VerifyIsRESOLVEDMIDInANRMatchString(mids.Value.AulPropTag[0]);
            this.VerifyIsUNRESOLVEDMIDInANREmptyString(mids.Value.AulPropTag[1]);
            this.VerifyIsAMBIGUOUSMIDInANRAmbiguousString(mids.Value.AulPropTag[2]);
            this.VerifyIsUNRESOLVEDMIDInANRNullString(mids.Value.AulPropTag[3]);
            this.VerifyIsUNRESOLVEDMIDInANRNotFound(mids.Value.AulPropTag[4]);
            this.VerifyIsResultOfANRProcessIsMID(mids);

            // Since if the whole ANR results are verified and the NspiResolveNamesW is invoked, 
            // it must take a set of string values in an 8-bit character set and perform ANR on those strings, 
            // this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1393,
                @"[In NspiResolveNamesW] The NspiResolveNamesW method takes a set of string values in the Unicode character set and performs ANR (as specified in section 3.1.4.7) on those strings.");

            // Since all the returned Minimal Entry IDs are verified according to the order of the input string array, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1426,
                @"[In NspiResolveNamesW] [Server Processing Rules: Upon receiving message NspiResolveNamesW, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] These Minimal Entry IDs are those that result from the ANR process, as specified in section 3.1.4.7, to the strings in the input parameter paWStr.");

            #endregion Capture
            #endregion

            #region Call NspiResolveNamesW method with a specific list of names and propertyTag value without null.

            PropertyTagArray_r propTagsWithProperties = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagObjectType
                }
            };

            this.Result = this.ProtocolAdatper.NspiResolveNamesW(reservedOfResolveNamesW, stat, propTagsWithProperties, wstrArray, out mids, out rowOfResolveNamesW);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiResolveNamesW should return Success!");
            Site.Assert.AreNotEqual<int>(0, rowOfResolveNamesW.Value.ARow.Length, "At least one address book object should be matched.");

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1395");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1395
            // The rowOfResolveNamesW which contains a PropertyRowSet_r structure is returned from the server. If it is not null, 
            // it illustrates that certain property values are returned for any valid Minimal Entry IDs identified by the ANR process.
            this.Site.CaptureRequirementIfIsNotNull(
                rowOfResolveNamesW,
                1395,
                @"[In NspiResolveNamesW] Certain property values are returned for any valid Minimal Entry IDs identified by the ANR process.");

            Site.Assert.AreEqual<uint>((uint)AulProp.PidTagEntryId, rowOfResolveNamesW.Value.ARow[0].LpProps[0].PropTag, "The first property returned should be PidTagEntryId. Now the returned property is {0}.", rowOfResolveNamesW.Value.ARow[0].LpProps[0].PropTag);
            Site.Assert.AreEqual<uint>((uint)AulProp.PidTagDisplayName, rowOfResolveNamesW.Value.ARow[0].LpProps[1].PropTag, "The second property returned should be PidTagDisplayName. Now the returned property is {0}.", rowOfResolveNamesW.Value.ARow[0].LpProps[1].PropTag);
            Site.Assert.AreEqual<uint>((uint)AulProp.PidTagObjectType, rowOfResolveNamesW.Value.ARow[0].LpProps[2].PropTag, "The third property returned should be PidTagObjectType. Now the returned property is {0}.", rowOfResolveNamesW.Value.ARow[0].LpProps[2].PropTag);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1399");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1399
            // Since the null value of pPropTags has been set in step "Call NspiResolveNamesW method with a specific list of names", 
            // and the above three asserts ensure that it is a reference to a PropertyTagArray_r value containing a list of the proptags of the columns that the client requests to be returned for each row.
            // So R1347 can be captured directly.
            this.Site.CaptureRequirement(
                1399,
                @"[In NspiResolveNamesW] pPropTags: The value NULL or a reference to a PropertyTagArray_r containing a list of the proptags of the columns that the client requests to be returned for each row returned.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the related requirements about string values returned from NspiResolveNames operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S03_TC03_ResolveNamesStringConversion()
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
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiResolveNames to require the string properties to be different with their native types.

            StringsArray_r strArray;
            strArray.CValues = 1;
            strArray.LppszA = new string[strArray.CValues];
            strArray.LppszA[0] = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            PropertyTagArray_r tags = new PropertyTagArray_r
            {
                CValues = 2,
                AulPropTag = new uint[]
                {
                    AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                    AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName)
                }
            };

            PropertyTagArray_r? mids;
            PropertyRowSet_r? rowOfResolveNames;

            this.Result = this.ProtocolAdatper.NspiResolveNames((uint)0, stat, tags, strArray, out mids, out rowOfResolveNames);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiResolveNames should return Success!");
            Site.Assert.IsNotNull(rowOfResolveNames, "PropertyRowSet_r value should not null. The row number is {0}.", rowOfResolveNames == null ? 0 : rowOfResolveNames.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowOfResolveNames.Value.ARow.Length, "At least one address book object should be matched.");

            PropertyRow_r rowValue = rowOfResolveNames.Value.ARow[0];

            string resolveddisplayNamePrintable = System.Text.UnicodeEncoding.UTF8.GetString(rowValue.LpProps[0].Value.LpszA);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1977: The value of PidTagAddressBookDisplayNamePrintable is {0}.", resolveddisplayNamePrintable);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1977
            // The field LpszA indicates an 8-bit character string value.
            Site.CaptureRequirementIfIsTrue(
                resolveddisplayNamePrintable.StartsWith(strArray.LppszA[0], StringComparison.Ordinal),
                1977,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiResolveNames method, String values can be returned in 8-bit character representation in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1976");

            string reslovedName = System.Text.UnicodeEncoding.Unicode.GetString(rowValue.LpProps[1].Value.LpszW);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1976
            // The field LpszW indicates a single Unicode string value.
            Site.CaptureRequirementIfAreEqual<string>(
                strArray.LppszA[0].ToLower(System.Globalization.CultureInfo.CurrentCulture),
                reslovedName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                1976,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiResolveNames method, String values can be returned in Unicode representation in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1938");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1938
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                rowValue.LpProps[0].PropTag,
                1938,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiResolveNames] If the native type of a property is PtypString and the client has requested that property with the type PtypString8, the server MUST convert the Unicode representation to an 8-bit character representation in the code page specified by the CodePage field of the pStat parameter prior to returning the value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1954");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1954
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName),
                rowValue.LpProps[1].PropTag,
                1954,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiResolveNames] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString, the server MUST convert the 8-bit character representation to a Unicode representation prior to returning the value.");

            #endregion

            #region Call NspiResolveNames to require the string properties to be the same as their native types.

            strArray.CValues = 1;
            strArray.LppszA = new string[strArray.CValues];
            strArray.LppszA[0] = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            tags = new PropertyTagArray_r
            {
                CValues = 2,
                AulPropTag = new uint[]
                {
                    (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                    (uint)AulProp.PidTagDisplayName
                }
            };

            this.Result = this.ProtocolAdatper.NspiResolveNames((uint)0, stat, tags, strArray, out mids, out rowOfResolveNames);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiResolveNames should return Success!");
            Site.Assert.IsNotNull(rowOfResolveNames, "PropertyRowSet_r value should not null. The row number is {0}.", rowOfResolveNames == null ? 0 : rowOfResolveNames.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowOfResolveNames.Value.ARow.Length, "At least one address book object should be matched.");

            rowValue = rowOfResolveNames.Value.ARow[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1946");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1938
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                rowValue.LpProps[0].PropTag,
                1946,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiResolveNames] If the native type of a property is PtypString and the client has requested that property with the type PtypString, the server MUST return the Unicode representation unmodified.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1962");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1962
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagDisplayName,
                rowValue.LpProps[1].PropTag,
                1962,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiResolveNames] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString8, the server MUST return the 8-bit character representation unmodified.");

            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the related requirements about string values returned from NspiResolveNamesW operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S03_TC04_ResolveNamesWStringConversion()
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
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiResolveNamesW to require the string properties to be different with their native type.
            uint reservedOfResolveNamesW = 0;
            WStringsArray_r wstrArray;
            wstrArray.CValues = 1;
            wstrArray.LppszW = new string[wstrArray.CValues];
            wstrArray.LppszW[0] = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 2,
                AulPropTag = new uint[]
                {
                    AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                    AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName)
                }
            };
            PropertyTagArray_r? mids;
            PropertyRowSet_r? rowOfResolveNamesW;

            this.Result = this.ProtocolAdatper.NspiResolveNamesW(reservedOfResolveNamesW, stat, propTags, wstrArray, out mids, out rowOfResolveNamesW);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiResolveNamesW should return Success!");
            Site.Assert.IsNotNull(rowOfResolveNamesW, "PropertyRowSet_r value should not null. The row number is {0}.", rowOfResolveNamesW == null ? 0 : rowOfResolveNamesW.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowOfResolveNamesW.Value.ARow.Length, "At least one address book object should be matched.");

            PropertyRow_r rowValue = rowOfResolveNamesW.Value.ARow[0];

            string resolvedDisplayNamePrintable = System.Text.UnicodeEncoding.UTF8.GetString(rowValue.LpProps[0].Value.LpszA);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1979: The value of PidTagAddressBookDisplayNamePrintable is {0}.", resolvedDisplayNamePrintable);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1979
            // The field LpszA indicates an 8-bit character string value.
            Site.CaptureRequirementIfIsTrue(
                resolvedDisplayNamePrintable.StartsWith(wstrArray.LppszW[0], StringComparison.Ordinal),
                1979,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiResolveNamesW method, String values can be returned in 8-bit character representation in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1978");

            string reslovedName = System.Text.UnicodeEncoding.Unicode.GetString(rowValue.LpProps[1].Value.LpszW);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1978
            // The field LpszW indicates a single Unicode string value.
            Site.CaptureRequirementIfAreEqual<string>(
                wstrArray.LppszW[0].ToLower(System.Globalization.CultureInfo.CurrentCulture),
                reslovedName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                1978,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiResolveNamesW method, String values can be returned in Unicode representation in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1939");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1939
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                rowValue.LpProps[0].PropTag,
                1939,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiResolveNamesW] If the native type of a property is PtypString and the client has requested that property with the type PtypString8, the server MUST convert the Unicode representation to an 8-bit character representation in the code page specified by the CodePage field of the pStat parameter prior to returning the value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1955");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1955
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName),
                rowValue.LpProps[1].PropTag,
                1955,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiResolveNamesW] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString, the server MUST convert the 8-bit character representation to a Unicode representation prior to returning the value.");

            #endregion

            #region Call NspiResolveNamesW to require the string properties to be the same as their native types.
            wstrArray.CValues = 1;
            wstrArray.LppszW = new string[wstrArray.CValues];
            wstrArray.LppszW[0] = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            propTags = new PropertyTagArray_r
            {
                CValues = 2,
                AulPropTag = new uint[]
                {
                    (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                    (uint)AulProp.PidTagDisplayName
                }
            };

            this.Result = this.ProtocolAdatper.NspiResolveNamesW(reservedOfResolveNamesW, stat, propTags, wstrArray, out mids, out rowOfResolveNamesW);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiResolveNamesW should return Success!");
            Site.Assert.IsNotNull(rowOfResolveNamesW, "PropertyRowSet_r value should not null. The row number is {0}.", rowOfResolveNamesW == null ? 0 : rowOfResolveNamesW.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowOfResolveNamesW.Value.ARow.Length, "At least one address book object should be matched.");

            rowValue = rowOfResolveNamesW.Value.ARow[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1946");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1947
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                rowValue.LpProps[0].PropTag,
                1947,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiResolveNamesW] If the native type of a property is PtypString and the client has requested that property with the type PtypString, the server MUST return the Unicode representation unmodified.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1963");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1963
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagDisplayName,
                rowValue.LpProps[1].PropTag,
                1963,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiResolveNamesW] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString8, the server MUST return the 8-bit character representation unmodified.");

            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        #region Private methods
        /// <summary>
        /// Check whether PropertyRowSet_r is not null.
        /// </summary>
        /// <param name="rowOfResolveNamesW">Contain the address book container rows that the server returns in response to the request.</param>
        private void VerifyPropertyRowSetIsNotNullForNspiResolveNamesW(PropertyRowSet_r? rowOfResolveNamesW)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1430");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1430
            // The rowOfResolveNamesW which contains a PropertyRowSet_r structure is returned from the server. If it is not null, 
            // it illustrates that the server must have constructed it.
            Site.CaptureRequirementIfIsNotNull(
                rowOfResolveNamesW,
                1430,
                @"[In NspiResolveNamesW] [Server Processing Rules: Upon receiving message NspiResolveNamesW, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] Subject to the prior constraints, the server MUST construct a PropertyRowSet_r to return to the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1405");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1405
            Site.CaptureRequirementIfIsNotNull(
                rowOfResolveNamesW,
                1405,
                @"[In NspiResolveNamesW] [ppRows: A reference to a PropertyRowSet_r structure (section 2.2.4), ] which contains the address book container rows that the server returns in response to the request. ");
        }

        /// <summary>
        /// Check whether PropertyRowSet_r is not null.
        /// </summary>
        /// <param name="rowOfResolveNames">Contain the address book container rows that the server returns in response to the request.</param>
        private void VerifyPropertyRowSetIsNotNullForNspiResolveNames(PropertyRowSet_r? rowOfResolveNames)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1379");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1379
            // The rowOfResolveNames which contains a PropertyRowSet_r structure is returned from the server. If it is not null, 
            // it illustrates that the server must have constructed it.
            Site.CaptureRequirementIfIsNotNull(
                rowOfResolveNames,
                1379,
                @"[In NspiResolveNames] [Server Processing Rules: Upon receiving message NspiResolveNames, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] Subject to the prior constraints, the server MUST construct a PropertyRowSet_r to return to the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1354");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1354
            Site.CaptureRequirementIfIsNotNull(
                rowOfResolveNames,
                1354,
                @"[In NspiResolveNames] [ppRows: A reference to a PropertyRowSet_r structure (section 2.2.4), ] which contains the address book container rows that the server returns in response to the request. ");
        }

        /// <summary>
        /// Check whether the result of the Ambiguous Name Resolution is Minimal Entry ID UNRESOLVED when the input string is not found.
        /// </summary>
        /// <param name="aulPropTag">A PropertyTag of PropertyTagArray_r. Its value should be 0, which means the ANR process is unable to map a string to any objects in the address book.</param>
        private void VerifyIsUNRESOLVEDMIDInANRNotFound(uint aulPropTag)
        {
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R88");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R88
                // According to the Open Specification, if the server is unable to map the string to any objects in the address book, 
                // the result of the ANR process is the Minimal Entry ID with the value MID_UNRESOLVE.
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)ANRMinEntryID.MID_UNRESOLVED,
                    aulPropTag,
                    88,
                    @"[In Ambiguous Name Resolution Minimal Entry IDs] MID_UNRESOLVED (0x00000000): The ANR process was unable to map a string to any objects in the address book.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R618");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R618
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)ANRMinEntryID.MID_UNRESOLVED,
                    aulPropTag,
                    618,
                    @"[In Ambiguous Name Resolution] If the server is unable to map the string to any objects in the address book, the result of the ANR process is the Minimal Entry ID with the value MID_UNRESOLVED.");
            }
        }

        /// <summary>
        /// Check whether the result of the Ambiguous Name Resolution is Minimal Entry ID UNRESOLVED when the input string is null.
        /// </summary>
        /// <param name="aulPropTag">A PropertyTag of PropertyTagArray_r. Its value should be 0, which means the ANR process is unable to map a string to any objects in the address book.</param>
        private void VerifyIsUNRESOLVEDMIDInANRNullString(uint aulPropTag)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R623");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R623
            // If the value that the client is requesting is set to NULL, and the returned Minimal Entry ID is MID_UNRESOLVED, MS-OXNSPI_R623 can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)ANRMinEntryID.MID_UNRESOLVED,
                aulPropTag,
                623,
                @"[In Ambiguous Name Resolution] The server MUST map the NULL string to the Minimal Entry ID MID_UNRESOLVED.");
        }

        /// <summary>
        /// Check whether the result of the Ambiguous Name Resolution is Minimal Entry ID UNRESOLVED when the input string is empty.
        /// </summary>
        /// <param name="aulPropTag">A PropertyTag of PropertyTagArray_r. Its value should be 0, which means the ANR process is unable to map a string to any objects in the address book.</param>
        private void VerifyIsUNRESOLVEDMIDInANREmptyString(uint aulPropTag)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R624");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R624
            // If the value that the client is requesting is set to an empty string, and the returned Minimal Entry ID is MID_UNRESOLVED, MS-OXNSPI_R624 can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)ANRMinEntryID.MID_UNRESOLVED,
                aulPropTag,
                624,
                @"[In Ambiguous Name Resolution] The server MUST map a zero-length string to the Minimal Entry ID MID_UNRESOLVED.");
        }

        /// <summary>
        /// Check whether the result of the Ambiguous Name Resolution is Minimal Entry ID RESOLVED when the input string is exactly matched.
        /// </summary>
        /// <param name="aulPropTag">A PropertyTag of PropertyTagArray_r. Its value should be 2, which means the ANR process maps a string to a single object in the address book.</param>
        private void VerifyIsRESOLVEDMIDInANRMatchString(uint aulPropTag)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R90");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R90
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)ANRMinEntryID.MID_RESOLVED,
                aulPropTag,
                90,
                @"[In Ambiguous Name Resolution Minimal Entry IDs] MID_RESOLVED (0x0000002): The ANR process mapped a string to a single object in the address book.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R622");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R622
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)ANRMinEntryID.MID_RESOLVED,
                aulPropTag,
                622,
                @"[In Ambiguous Name Resolution] If the server is able to map the string to exactly one object in the address book,
                the result of the ANR process is the Minimal Entry ID with the value MID_RESOLVED.");
        }

        /// <summary>
        /// Check whether the result of the Ambiguous Name Resolution is Minimal Entry ID RESOLVED when the input string is matched to multiple objects.
        /// </summary>
        /// <param name="aulPropTag">A PropertyTag of PropertyTagArray_r. Its value should be 1, which means the ANR process maps a string to multiple objects in the address book.</param>
        private void VerifyIsAMBIGUOUSMIDInANRAmbiguousString(uint aulPropTag)
        {
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R89");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R89
                // According to the Open Specification, if the server is able to map the string to more than one object in the address book, 
                // the result of the ANR process is the Minimal Entry ID with the value MID_AMBIGUOUS.
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)ANRMinEntryID.MID_AMBIGUOUS,
                    aulPropTag,
                    89,
                    @"[In Ambiguous Name Resolution Minimal Entry IDs] MID_AMBIGUOUS (0x0000001):  The ANR process mapped a string to multiple objects in the address book.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R620");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R620
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)ANRMinEntryID.MID_AMBIGUOUS,
                    aulPropTag,
                    620,
                    @"[In Ambiguous Name Resolution] If the server is able to map the string to more than one object in the address book, the result of the ANR process is the Minimal Entry ID with the value MID_AMBIGUOUS.");
            }
        }

        /// <summary>
        /// Check whether the specific result of the Ambiguous Name Resolution process is Minimal Entry ID RESOLVED.
        /// </summary>
        /// <param name="mids">A PropertyTagArray_r value which contains a list of Minimal Entry IDs that match the array of strings.</param>
        private void VerifyIsResultOfANRProcessIsMID(PropertyTagArray_r? mids)
        {
            for (int i = 0; i < mids.Value.AulPropTag.Length; i++)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R616, the property tag {0} of minimal entry id is {1}", i, mids.Value.AulPropTag[i]);

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R616
                bool isVerifiedR616 = mids.Value.AulPropTag[i] == (uint)ANRMinEntryID.MID_AMBIGUOUS ||
                                      mids.Value.AulPropTag[i] == (uint)ANRMinEntryID.MID_RESOLVED ||
                                      mids.Value.AulPropTag[i] == (uint)ANRMinEntryID.MID_UNRESOLVED;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR616,
                    616,
                    @"[In Ambiguous Name Resolution] The specific result of an ANR process is a Minimal Entry ID.");
            }
        }

        #endregion
    }
}