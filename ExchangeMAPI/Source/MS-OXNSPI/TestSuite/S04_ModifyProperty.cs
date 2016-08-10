namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class contains all the test cases designed to test the server behavior for the NSPI calls related to modify property of Address Book Object.
    /// </summary>
    [TestClass]
    public class S04_ModifyProperty : TestSuiteBase
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
        /// This test case is designed to verify the requirements related to NspiModProps operation returning success.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S04_TC01_ModPropsSuccess()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between a client and the server.
            uint flags = 0;
            STAT stat = new STAT();
            stat.InitiateStat();
            FlatUID_r guid = new FlatUID_r();
            guid.Ab = new byte[16];
            FlatUID_r? serverGuid = guid;
            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");

            #endregion

            #region Call NspiQueryRows to get the DN of specified user.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fSkipObjects;
            uint tableCount = 0;
            uint[] table = null;
            uint count = Constants.QueryRowsRequestedRowNumber;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r();
            propTagsInstance.CValues = 3;
            propTagsInstance.AulPropTag = new uint[3]
            {
                (uint)AulProp.PidTagEntryId,
                (uint)AulProp.PidTagDisplayName,
                (uint)AulProp.PidTagDisplayType,
            };
            PropertyTagArray_r? propTags = propTagsInstance;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return success!");

            string userName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
            string userESSDN = string.Empty;

            for (int i = 0; i < rowsOfQueryRows.Value.CRows; i++)
            {
                string name = System.Text.Encoding.Default.GetString(rowsOfQueryRows.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(userName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    PermanentEntryID administratorEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.Bin.Lpb);
                    userESSDN = administratorEntryID.DistinguishedName;

                    break;
                }
            }
            #endregion

            #region Call NspiDNToMId to get the MIDs of specified user.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r();
            names.LppszA = new string[]
            {
                userESSDN
            };
            names.CValues = (uint)names.LppszA.Length;
            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            #endregion

            #region Call NspiGetMatches to get the specific PidTagAddressBookX509Certificate property to be modified.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = 1;
            stat.CurrentRec = mids.Value.AulPropTag[0];
            Restriction_r? filter = null;
            PropertyTagArray_r propTags1 = new PropertyTagArray_r();
            propTags1.CValues = 2;
            propTags1.AulPropTag = new uint[2]
            {
                (uint)AulProp.PidTagAddressBookX509Certificate,
                (uint)AulProp.PidTagUserX509Certificate
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags1;

            // Set value for propNameOfGetMatches.
            PropertyName_r? propNameOfGetMatches = null;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return success!");
            Site.Assert.IsNotNull(outMIds.Value, "The Minimal Entry IDs returned successfully.");
            #endregion

            #region Call NspiModProps method with specific PidTagAddressBookX509Certificate property value.
            uint reservedOfModProps = 1;
            BinaryArray_r emptyValue = new BinaryArray_r();
            PropertyRow_r rowOfModProps = new PropertyRow_r();
            rowOfModProps.LpProps = new PropertyValue_r[2];
            rowOfModProps.LpProps[0].PropTag = (uint)AulProp.PidTagAddressBookX509Certificate;
            rowOfModProps.LpProps[0].Value.MVbin = emptyValue;
            rowOfModProps.LpProps[1].PropTag = (uint)AulProp.PidTagUserX509Certificate;
            rowOfModProps.LpProps[1].Value.MVbin = emptyValue;

            PropertyTagArray_r instanceOfModProps = new PropertyTagArray_r();
            instanceOfModProps.CValues = 2;
            instanceOfModProps.AulPropTag = new uint[2]
            {
                (uint)AulProp.PidTagAddressBookX509Certificate,
                (uint)AulProp.PidTagUserX509Certificate
            };
            PropertyTagArray_r? propTagsOfModProps = instanceOfModProps;

            this.Result = this.ProtocolAdatper.NspiModProps(reservedOfModProps, stat, propTagsOfModProps, rowOfModProps);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1305");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1305
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1305,
                @"[In NspiModProps] [Server Processing Rules: Upon receiving message NspiModProps, the server MUST process the data from the message subject to the following constraints:] [constraint 12] If no other return values have been specified by these constraints [constraints 1-11], the server MUST return the return value ""Success"".");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5309,
                @"[In PidTagAddressBookX509Certificate] Property ID: 0x8C6A.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5310,
                @"[In PidTagAddressBookX509Certificate] Data type: PtypMultipleBinary, 0x1102.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                8875,
                @"[In PidTagUserX509Certificate] Property ID: 0x3A70.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                8876,
                @"[In PidTagUserX509Certificate] Data type: PtypMultipleBinary, 0x1102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1267");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1267
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1267,
                @"[In NspiModProps] The NspiModProps method is used to modify the properties of an object in the address book.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1268");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1268
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1268,
                @"[In NspiModProps] This protocol supports the PidTagUserX509Certificate ([MS-OXPROPS] section 2.1044) and PidTagAddressBookX509Certificate ([MS-OXPROPS] section 2.566) properties.");

            bool isR1289Verified = reservedOfModProps != 0 && ErrorCodeValue.Success == this.Result;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1289");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1289
            Site.CaptureRequirementIfIsTrue(
                isR1289Verified,
                1289,
                @"[In NspiModProps] [Server Processing Rules: Upon receiving message NspiModProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] If the Reserved input parameter contains any value other than 0, the server MUST ignore the value.");
            #endregion
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiModLinkAtt operation returning success with PidTagAddressBookMember.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S04_TC02_ModLinkAttSuccessWithPidTagAddressBookMember()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();
            bool isR2003009Enabled = Common.IsRequirementEnabled(2003009, this.Site);

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

            #region Call NspiGetMatches method to get the valid Minimal Entry IDs and rows.
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
            string dlistName = Common.GetConfigurationPropertyValue("DistributionListName", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(dlistName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(dlistName + "\0");
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

            #region Call NspiModLinkAtt method with the specified PidTagAddressBookMember value.
            uint flagsOfModLinkAtt = 0; // A value which does not contain fDelete flag (0x1).
            uint propTagOfModLinkAtt = (uint)AulProp.PidTagAddressBookMember;
            uint midOfModLinkAtt = 0;
            BinaryArray_r entryId = new BinaryArray_r
            {
                CValues = 1,
                Lpbin = new Binary_r[1]
            };

            int i;
            string name = string.Empty;
            for (i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(dlistName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    // The current Address Book object is a distribution list.
                    // Save the distribution list location.
                    midOfModLinkAtt = outMIds.Value.AulPropTag[i];
                    this.MidToBeModified = midOfModLinkAtt;
                }
                else if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(memberName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    // Save the EntryID of this user which will be added as a member of the distribution list.
                    entryId.Lpbin[0] = rowsOfGetMatches.Value.ARow[i].LpProps[0].Value.Bin;
                    this.EntryIdToBeDeleted = entryId;
                }

                if (this.MidToBeModified != 0 && this.EntryIdToBeDeleted.Lpbin[0].Cb != 0)
                {
                    break;
                }
            }

            Site.Assert.AreEqual<uint>(0x87, entryId.Lpbin[0].Lpb[0], "Ephemeral Entry ID's ID Type should be 0x87.");
            this.IsEphemeralEntryID = true;

            // Add the property value.
            ErrorCodeValue result1;
            if (!isR2003009Enabled)
            {
                result1 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1340");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1340
                Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                    ErrorCodeValue.Success,
                    result1,
                    1340,
                    @"[In NspiModLinkAtt] [Server Processing Rules: Upon receiving message NspiModLinkAtt, the server MUST process the data from the message subject to the following constraints:] [Constraint 9] If no other return values have been specified by these constraints [constraints 1-8], the server MUST return the return value ""Success"" (0x00000000).");
            }
            else
            {
                result1 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId, false);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R2003009");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R2003009
                Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                    ErrorCodeValue.GeneralFailure,
                    result1,
                    2003009,
                    @"[In Appendix A: Product Behavior] Implementation does return ""GeneralFailure"" when modify either the PidTagAddressBookMember property or the PidTagAddressBookPublicDelegates property of any objects in the address book. <6> Section 3.1.4.1.15:  Exchange 2013 and Exchange 2016 return ""GeneralFailure"" (0x80004005) when modification of either the PidTagAddressBookMember property ([MS-OXOABK] section 2.2.6.1) or the PidTagAddressBookPublicDelegates property ([MS-OXOABK] section 2.2.5.5) is attempted.");
            }

            this.IsRequireToDeleteAddressBookMember = true;
            #endregion

            #region Call NspiGetMatches to check the NspiModLinkAtt result.
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            #region Capture code
            // PidTagAddressBookMember proptag after adding value.
            uint addressBookMemberTagAfterAdd = 0;
            for (i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(dlistName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    // After adding value this value should not be 0x8009000a.
                    addressBookMemberTagAfterAdd = rowsOfGetMatches.Value.ARow[i].LpProps[2].PropTag;

                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1335, the property tag of address book member after adding value is {0}",
                addressBookMemberTagAfterAdd);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1335
            // The proptag of PidTagAddressBookMember property is 0x8009101e as defined in [MS-OXPROPS]. If the last four bytes is 0x000a when being returned, it means that the property has no value.
            // Since the property has no value before the NspiModLinkAtt method is called and has value after that, this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                addressBookMemberTagAfterAdd != 0x8009000a,
                1335,
                @"[In NspiModLinkAtt] [Server Processing Rules: Upon receiving message NspiModLinkAtt, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] If the input parameter dwFlags does not contain the bit value fDelete, the server MUST add all values specified by the input parameter lpEntryIDs to the property specified by ulPropTag for the object specified by the input parameter dwMId.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R129, the property tag of address book member after adding value is {0}",
                addressBookMemberTagAfterAdd);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R129
            // Since the property has no value before the NspiModLinkAtt method is called and has value after that, this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                addressBookMemberTagAfterAdd != 0x8009000a,
                129,
                @"[In NspiModLinkAtt Flags] If the fDelete flag is not set, the server adds values when modifying.");

            // MS-OXNSPI_R129 has already verified that modifying property PidTagAddressBookMember succeeds, so MS-OXNSPI_R1893 can be verified directly here.
            this.Site.CaptureRequirement(
                1893,
                @"[In NspiModLinkAtt] This protocol supports modifying the value of the PidTagAddressBookMember ([MS-OXPROPS] section 2.541) property of an address book object with display type DT_DISTLIST.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1306");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1306
            this.Site.CaptureRequirement(
                1306,
                @"[In NspiModLinkAtt] The NspiModLinkAtt method modifies the values of a specific property of a specific row in the address book.");

            #endregion
            #endregion

            #region Call NspiModLinkAtt with fDelete flag to delete the specified value.
            flagsOfModLinkAtt = (uint)NspiModLinkAtFlag.fDelete;
            ErrorCodeValue result2;
            if (!isR2003009Enabled)
            {
                result2 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result2, "NspiModLinkAtt method should return Success.");
            }
            else
            {
                result2 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId, false);
            }

            this.IsRequireToDeleteAddressBookMember = false;
            #endregion

            #region Call NspiGetMatches to check the NspiModLinkAtt delete result.
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            #region Capture code
            bool isDeleteSuccess = false;
            for (i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(dlistName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    // 0x8009000A means that no member exists in distribute list, so the added member has been deleted.
                    if (rowsOfGetMatches.Value.ARow[i].LpProps[2].PropTag == 0x8009000A)
                    {
                        isDeleteSuccess = true;
                        break;
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1332: the returned PidTagAddressBookMember proptag value of the address book object named {0} is {1}", name, rowsOfGetMatches.Value.ARow[i].LpProps[2].PropTag);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1332
            // 0x8009000A means that no member exists in distribute list, so the added member has been deleted.
            Site.CaptureRequirementIfIsTrue(
                isDeleteSuccess,
                1332,
                @"[In NspiModLinkAtt] [Server Processing Rules: Upon receiving message NspiModLinkAtt, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] If the input parameter dwFlags contains the bit value fDelete, the server MUST remove all values specified by the input parameter lpEntryIDs from the property specified by ulPropTag for the object specified by input parameter dwMId.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R128: the returned PidTagAddressBookMember property tag value of the address book object named {0} is {1}", name, rowsOfGetMatches.Value.ARow[i].LpProps[2].PropTag);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R128
            // 0x8009000A means that no member exists in distribute list, so the added member has been deleted.
            Site.CaptureRequirementIfIsTrue(
                isDeleteSuccess,
                128,
                @"[In NspiModLinkAtt Flags] fDelete (0x00000001): Specifies that the server is to remove values when modifying.");

            #endregion
            #endregion

            #region Call NspiModLinkAtt to add value with the changed Display Type in EphemeralEntryID.
            // If the original display type is DT_MAILUSER (0x00), change it to DT_PRIVATE_DISTLIST (0x05), otherwise to DT_MAILUSER (0x00), to see if the server ignores this field.
            // The Display Type begins with the 25th position in the Entry ID structure.
            if (entryId.Lpbin[0].Lpb[24] == 0x00)
            {
                entryId.Lpbin[0].Lpb[24] = 0x05;
            }
            else
            {
                entryId.Lpbin[0].Lpb[24] = 0x00;
            }

            ErrorCodeValue result3;
            if (!isR2003009Enabled)
            {
                result3 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result3, "NspiModLinkAtt method should return Success.");
            }
            else
            {
                result3 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId, false);
            }

            this.IsRequireToDeleteAddressBookMember = true;
            this.EntryIdToBeDeleted = entryId;
            #endregion

            #region Call NspiModLinkAtt to delete value with the changed Display Type in EphemeralEntryID.
            flagsOfModLinkAtt = (uint)NspiModLinkAtFlag.fDelete;
            ErrorCodeValue result4;
            if (!isR2003009Enabled)
            {
                result4 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result4, "NspiModLinkAtt method should return Success.");
            }
            else
            {
                result4 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId, false);
            }

            this.IsRequireToDeleteAddressBookMember = false;

            #region Capture code
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1665, the value of result1 is {0}, the value of result2 is {1}, the value of result3 is {2}, the value of result4 is {3}",
                result1,
                result2,
                result3,
                result4);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1665
            // Since the server returns the same value when the display type field is set to different values in the input parameter, whenever the property value is to be added (result1 and result3) or to be deleted (result2 and result4), this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                result1 == result3 && result2 == result4,
                1665,
                @"[In EphemeralEntryID] If this field[Display Type ] is set to different values, the server will return the same value.");

            #endregion
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the related requirements about NspiModLinkAtt operation with different display types of PermanentEntryID.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S04_TC03_ModLinkAttSuccessWithDifferentDisplayTypePermanentEntryID()
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

            #region Call NspiQueryRows method to get a set of valid rows used to get matched entry ID as the input parameter of NspiModLinkAtt method.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fSkipObjects; // Since the flag fEphID (0x2) is not set, the returned Entry ID is Permanent Entry ID.
            uint tableCount = 0;
            uint[] table = null;
            uint count = Constants.QueryRowsRequestedRowNumber;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 4,
                AulPropTag = new uint[4]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagAddressBookMember
                }
            };
            PropertyTagArray_r? propTags = propTagsInstance;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.IsTrue(this.Result == ErrorCodeValue.Success || this.Result == ErrorCodeValue.ErrorsReturned, "NspiQueryRows should return Success or ErrorsReturned (which just specify some properties in the result have no value)!");
            #endregion

            #region Call NspiModLinkAtt to add the specified PidTagAddressBookMember value.
            uint flagsOfModLinkAtt = 0; // A value which does not contain fDelete flag (0x1).
            uint propTagOfModLinkAtt = (uint)AulProp.PidTagAddressBookMember;
            uint midOfModLinkAtt = 0;
            BinaryArray_r entryId = new BinaryArray_r
            {
                CValues = 1,
                Lpbin = new Binary_r[1]
            };

            string dlistName = Common.GetConfigurationPropertyValue("DistributionListName", this.Site);
            string memberName = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            for (int i = 0; i < rowsOfQueryRows.Value.CRows; i++)
            {
                string name = System.Text.Encoding.Default.GetString(rowsOfQueryRows.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(dlistName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    PermanentEntryID entryID = AdapterHelper.ParsePermanentEntryIDFromBytes(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.Bin.Lpb);

                    // To get MId of the DN in Entry Id.
                    #region NspiDNToMId
                    uint reserved = 0;
                    StringsArray_r names = new StringsArray_r
                    {
                        CValues = 1,
                        LppszA = new string[1]
                    };
                    names.LppszA[0] = entryID.DistinguishedName;
                    PropertyTagArray_r? mids;
                    this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
                    #endregion

                    midOfModLinkAtt = mids.Value.AulPropTag[0];
                    this.MidToBeModified = midOfModLinkAtt;
                }
                else if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(memberName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    entryId.Lpbin[0] = rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.Bin;
                    this.EntryIdToBeDeleted = entryId;
                }

                if (midOfModLinkAtt != 0 && entryId.Lpbin[0].Cb != 0)
                {
                    break;
                }
            }

            Site.Assert.AreEqual<uint>(0x00, entryId.Lpbin[0].Lpb[0], "Permanent Entry ID's ID Type should be 0x00.");

            // Add the specified PidTagAddressBookMember value.
            ErrorCodeValue result1 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result1, "NspiModLinkAtt method should return Success.");
            this.IsRequireToDeleteAddressBookMember = true;
            #endregion

            #region Call NspiModLinkAtt to delete the specified PidTagAddressBookMember value.
            flagsOfModLinkAtt = (uint)NspiModLinkAtFlag.fDelete;
            ErrorCodeValue result2 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result2, "NspiModLinkAtt method should return Success.");
            this.IsRequireToDeleteAddressBookMember = false;
            #endregion

            #region Call NspiModLinkAtt to add the specified PidTagAddressBookMember value with the changed display type in PermanentEntryID.
            flagsOfModLinkAtt = 0; // A value which does not contain fDelete flag (0x1).

            // If the original display type is DT_MAILUSER (0x00), change it to DT_PRIVATE_DISTLIST (0x05), otherwise to DT_MAILUSER (0x00), to see if the server ignores this field.
            // The Display Type begins with the 24th position in the Entry ID structure.
            if (entryId.Lpbin[0].Lpb[24] == 0x00)
            {
                entryId.Lpbin[0].Lpb[24] = 0x05;
            }
            else
            {
                entryId.Lpbin[0].Lpb[24] = 0x00;
            }

            ErrorCodeValue result3 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result3, "NspiModLinkAtt method should return Success.");
            this.IsRequireToDeleteAddressBookMember = true;
            #endregion

            #region Call NspiModLinkAtt to delete the specified PidTagAddressBookMember value with the changed display type in PermanentEntryID.
            flagsOfModLinkAtt = (uint)NspiModLinkAtFlag.fDelete;
            ErrorCodeValue result4 = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result4, "NspiModLinkAtt method should return Success.");
            this.IsRequireToDeleteAddressBookMember = false;

            #region Capture code
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1666:�Call NspiModLinkAtt to add address book member�{0} and�delete�it�{1}, then�add address book member with different display type�{2} and�delete�it�{3}.",
                result1.ToString(),
                result2.ToString(),
                result3.ToString(),
                result4.ToString());

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1666
            // Since the server returns the same value when the display type field is set to different values in the input parameter, whenever the property value is to be added (result1 and result3) or to be deleted (result2 and result4), this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                result1 == result3 && result2 == result4,
                1666,
                @"[In PermanentEntryID] If this field [Display Type String ] is set to different values, the server will return the same result.");

            #endregion

            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiModLinkAtt operation returning success with PidTagAddressBookPublicDelegates.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S04_TC04_ModLinkAttSuccessWithPidTagAddressBookPublicDelegates()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();
            bool isR2003009Enabled = Common.IsRequirementEnabled(2003009, this.Site);

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;
            STAT stat = new STAT();
            stat.InitiateStat();
            FlatUID_r guid = new FlatUID_r();
            guid.Ab = new byte[16];
            FlatUID_r? serverGuid = guid;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");

            #endregion

            #region Call NspiGetMatches method to get valid Minimal Entry IDs and rows.
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
            string userName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(userName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(userName + "\0");
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
            string administratorName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(administratorName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(administratorName + "\0");
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

            #region Call NspiModLinkAtt method to add the specified PidTagAddressBookPublicDelegates value.
            uint flag1 = 0x00;
            uint propTagOfModLinkAtt = (uint)AulProp.PidTagAddressBookPublicDelegates;
            uint midOfModLinkAtt = 0;
            BinaryArray_r entryId = new BinaryArray_r
            {
                CValues = 1,
                Lpbin = new Binary_r[1]
            };

            // Get user name
            string name = string.Empty;
            int i = 0;
            for (i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                name = System.Text.Encoding.UTF8.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(userName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    midOfModLinkAtt = outMIds.Value.AulPropTag[i];
                    this.MidToBeModified = midOfModLinkAtt;
                }
                else if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(administratorName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    entryId.Lpbin[0] = rowsOfGetMatches.Value.ARow[i].LpProps[0].Value.Bin;
                    this.EntryIdToBeDeleted = entryId;
                }

                if (this.MidToBeModified != 0 && this.EntryIdToBeDeleted.Lpbin[0].Cb != 0)
                {
                    break;
                }
            }

            ErrorCodeValue flag1Result;
            if (!isR2003009Enabled)
            {
                flag1Result = this.ProtocolAdatper.NspiModLinkAtt(flag1, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, flag1Result, "NspiModLinkAtt method should return Success.");
            }
            else
            {
                flag1Result = this.ProtocolAdatper.NspiModLinkAtt(flag1, propTagOfModLinkAtt, midOfModLinkAtt, entryId, false);
            }

            this.IsRequireToDeleteAddressBookPublicDelegate = true;
            #endregion

            #region Call NspiGetMatches to check the NspiModLinkAtt result.
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            // PidTagAddressBookPublicDelegates proptag after adding value.
            uint addressBookPublicDelegatesTagAfterAdd = 0;
            for (i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(userName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    // After adding value this value should not be 0x8015000a
                    addressBookPublicDelegatesTagAfterAdd = rowsOfGetMatches.Value.ARow[i].LpProps[2].PropTag;

                    break;
                }
            }

            #endregion

            #region Call NspiModLinkAtt method to delete the specified PidTagAddressBookPublicDelegates value.
            uint flagsOfModLinkAtt = (uint)NspiModLinkAtFlag.fDelete;
            if (!isR2003009Enabled)
            {
                this.Result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiModLinkAtt method should return Success.");
            }
            else
            {
                this.Result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId, false);
            }

            this.IsRequireToDeleteAddressBookPublicDelegate = false;
            #endregion

            #region Call NspiGetMatches to check the NspiModLinkAtt result.
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            #region Capture code
            // PidTagAddressBookPublicDelegates proptag after deleting value.
            uint addressBookPublicDelegatesTagAfterDelete = 0;
            for (i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                name = System.Text.Encoding.Default.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(userName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    // After deleting value this value should be 0x8015000a
                    addressBookPublicDelegatesTagAfterDelete = rowsOfGetMatches.Value.ARow[i].LpProps[2].PropTag;

                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXNSPI_R2002, the property tag of address book delegate after adding value is {0}, and the property tag of it after deleting value is {1}",
                addressBookPublicDelegatesTagAfterAdd,
                addressBookPublicDelegatesTagAfterDelete);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R2002
            // The proptag of PidTagAddressBookPublicDelegates property is 0x8015101e as defined in [MS-OXPROPS]. If the last four bytes is 0x000a when returned, it means the property has no value.
            // Since the property has no value before the NspiModLinkAtt method is called to add value, and has value after that, then has no value after deleting the value, this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                addressBookPublicDelegatesTagAfterAdd != 0x8015000a && addressBookPublicDelegatesTagAfterDelete == 0x8015000a,
                2002,
                @"[In NspiModLinkAtt] This protocol supports modifying the value of the PidTagAddressBookPublicDelegates ([MS-OXPROPS] section 2.557) property of an address book object with display type DT_MAILUSER.");

            #endregion
            #endregion

            #region Call NspiModLinkAtt method twice with different flag value.
            flag1 = 0xff;
            uint flag2 = 0xfe;

            if (!isR2003009Enabled)
            {
                flag1Result = this.ProtocolAdatper.NspiModLinkAtt(flag1, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, flag1Result, "NspiModLinkAtt method should return Success.");
            }
            else
            {
                flag1Result = this.ProtocolAdatper.NspiModLinkAtt(flag1, propTagOfModLinkAtt, midOfModLinkAtt, entryId, false);
            }

            this.IsRequireToDeleteAddressBookPublicDelegate = true;

            ErrorCodeValue flag2Result;
            if (!isR2003009Enabled)
            {
                flag2Result = this.ProtocolAdatper.NspiModLinkAtt(flag2, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, flag2Result, "NspiModLinkAtt method should return Success.");
            }
            else
            {
                flag2Result = this.ProtocolAdatper.NspiModLinkAtt(flag2, propTagOfModLinkAtt, midOfModLinkAtt, entryId, false);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1927");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1927
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                flag1Result,
                flag2Result,
                1927,
                @"[In NspiModLinkAtt] If dwFlags is set to different values other than the bit flag fDelete, server will return the same result.");
            #endregion

            #region Call NspiModLinkAtt method to delete the specified PidTagAddressBookPublicDelegates value.
            if (!isR2003009Enabled)
            {
                this.Result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiModLinkAtt method should return Success.");
            }
            else
            {
                this.Result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, midOfModLinkAtt, entryId, false);
            }

            this.IsRequireToDeleteAddressBookPublicDelegate = false;
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }
    }
}