namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class contains all the test cases designed to test the server behavior for the NSPI calls related to obtaining the detailed information of Address Book Object.
    /// </summary>
    [TestClass]
    public class S02_ObtainDetailsInfo : TestSuiteBase
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
        /// This test case is designed to verify the requirements related to NspiUpdateStat operation with different CurrentRec.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC01_UpdateStatSuccessWithDifferentCurrentRec()
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

            #region Call NspiUpdateStat with the CurrentRec field of the input parameter pStat not set to MID_CURRENT.
            stat.InitiateStat();
            stat.CurrentRec = (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE;
            uint reserved = 0;
            stat.Delta = 0;
            int? delta = 2;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R542");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R542
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                542,
                @"[In Absolute Positioning] [step 4] The server MUST support the special Minimal Entry ID MID_BEGINNING_OF_TABLE, as specified in section 2.2.1.8.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R800");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R800
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                800,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 12] If no error condition has been specified by the previous constraints, the server MUST return ""Success"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R798");

            uint objectNum = this.SutControlAdapter.GetNumberOfAddressBookObject();

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R798
            Site.CaptureRequirementIfAreEqual<uint>(
                objectNum,
                stat.TotalRecs,
                798,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 10] The server MUST set the TotalRecs field of the parameter pStat to the number of rows in the current address book container according to section 3.1.4.5.2.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R332");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R332
            Site.CaptureRequirementIfAreEqual<uint>(
                objectNum,
                stat.TotalRecs,
                332,
                @"[In STAT] [TotalRecs] The server sets this field to specify the total number of rows in the table.");

            #endregion
            #endregion

            #region Call NspiUpdateStat with the CurrentRec field of the input parameter pStat set to MID_CURRENT.
            stat.InitiateStat();
            stat.CurrentRec = (uint)MinimalEntryID.MID_CURRENT;
            reserved = 0;
            delta = 2;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R575");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R575
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                stat.TotalRecs,
                575,
                @"[In Fractional Positioning] [step 3] The server reports this [the number of objects] in the TotalRecs field of the STAT structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1742");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1742
            // CurrentRec in stat is set to MID_Current, if server returns success, it means the MID_Current is valid in method NspiUpdateStat and related requirements are verified.
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1742,
                @"[In Positioning Minimal Entry IDs] [MID_CURRENT] This Minimal Entry ID is valid in the NspiUpdateStat method.");

            #endregion
            #endregion

            #region Call NspiUpdateStat with the CurrentRec field of the input parameter pStat set to MID_END_OF_TABLE.
            delta = 0;
            stat.CurrentRec = (uint)MinimalEntryID.MID_END_OF_TABLE;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1758");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1758
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1758,
                @"[In Absolute Positioning] The server MUST support the special Minimal Entry ID MID_END_OF_TABLE, as specified in section 2.2.8.");

            #endregion Capture
            #endregion

            #region Call NspiUpdateStat to move CurrentRec to be before the first row of the table.
            stat.InitiateStat();
            stat.CurrentRec = (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE;
            reserved = 0;
            stat.Delta = -1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R558");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R558
            // According to the Open Specification, if applying the Delta would move the Current Position to be before the first row of the table, 
            // the server sets the Current Position to the first row of the table. So if stat.NumPos equals 0, it indicates the numeric row is 0-based.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x0,
                stat.NumPos,
                558,
                @"[In Absolute Positioning] This numeric row [the numeric row of the Current Position in the sorted table] is 0-based.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R561");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R561
            // Now the current position is pointed to be the first row of the table, and the numeric row is 0-based. So if stat.NumPos equals 0, 
            // it indicates server reports the current position in the NumPos field of the STAT structure.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x0,
                stat.NumPos,
                561,
                @"[In Absolute Positioning] The server reports this [the Numeric Position of the Current Position of the table] 
                in the NumPos field of the STAT structure.");

            #endregion Capture
            #endregion

            #region Call NspiUpdateStat to move CurrentRec to be after the end of the table.
            stat.InitiateStat();
            stat.CurrentRec = (uint)MinimalEntryID.MID_END_OF_TABLE;
            reserved = 0;
            stat.Delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1704");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1704
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_END_OF_TABLE,
                stat.CurrentRec,
                1704,
                @"[In Absolute Positioning] [step 8] If applying the Delta as described in step 6 would move the Current Position to
                be after the end of the table, the server sets the CurrentRec to the value MID_END_OF_TABLE.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiUpdateStat with values of some fields that are ignored by server.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC02_UpdateStatIgnoreSomeFlags()
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

            #region Call NspiUpdateStat with a specified Reserved value.
            uint reserved = 0;
            int? delta = 0;
            stat.Delta = 2;
            ErrorCodeValue result1 = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result1, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiUpdateStat with a different value from that in the above step.
            reserved = 1;
            stat.Delta = 2;
            ErrorCodeValue result2 = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result2, "NspiUpdateStat should return Success!");

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1676");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1676
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                result1,
                result2,
                1676,
                @"[In NspiUpdateStat] If this field[Reserved] is set to different values, the server will return the same result.");
            #endregion
            #endregion

            #region Call NspiUpdateStat with plDelta set to 1.
            int? delta1 = 1;
            ErrorCodeValue result3 = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result3, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiUpdateStat with plDelta set to 2.
            int? delta2 = 2;
            ErrorCodeValue result4 = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta2);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result4, "NspiUpdateStat should return Success!");

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1872");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1872
            this.Site.CaptureRequirementIfAreEqual<int?>(
                delta1,
                delta2,
                1872,
                @"[In NspiUpdateStat] If this field [plDelta] is set to different values rather than NULL, server will return the same result.");

            #endregion
            #endregion

            #region Call NspiUpdateStat with plDelta set to null.
            delta = null;
            ErrorCodeValue result5 = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result5, "NspiUpdateStat should return Success!");

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1868");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1868
            this.Site.CaptureRequirementIfIsNull(
                delta,
                1868,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] If the input parameter plDelta is null, the server MUST set the output parameter plDelta to null.");
            #endregion

            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiUpdateStat operation with different Delta.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC03_UpdateStatSuccessWithDifferentDelta()
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

            #region Call NspiUpdateStat to update STAT block to move forward 1 row from 0.
            uint reserved = 0;
            int? delta = 0;
            stat.Delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            // Save the CurrentRec of first row.
            uint frontPosition = (uint)stat.CurrentRec;
            #endregion

            #region Call NspiUpdateStat to update STAT block to move forward 3 rows from 0.
            stat.InitiateStat();
            stat.Delta = 3;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            // Save the CurrentRec of the third row.
            uint laterPosition = (uint)stat.CurrentRec;
            #endregion

            #region Call NspiUpdateStat to update STAT block to move forward 2 rows from the first row position.
            stat.CurrentRec = frontPosition;
            stat.Delta = 2;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            // Save TotalRecs value.
            uint totalRecs3 = stat.TotalRecs;

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R333", "The total number of rows in the table is {0}.", totalRecs3);

            uint expectedValue = this.SutControlAdapter.GetNumberOfAddressBookObject();

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R333
            // If the total rows number of a table equals the value configured in the ptfconfig file, it specifies that the value of TotalRecs is accurate.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                expectedValue,
                totalRecs3,
                333,
                @"[In STAT] [TotalRecs] Unlike the NumPos field, the server MUST report this number accurately; an approximation is insufficient.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R795. The value of CurrentRec is {0}. The value of delta is {1}.", stat.CurrentRec, delta);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R795
            // The final position should equal laterPosition and delta should equal 2
            Site.CaptureRequirementIfIsTrue(
                (stat.CurrentRec == laterPosition) && (delta == 2),
                795,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] If the input parameter plDelta is not null, the server MUST set it [plDelta] to the actual number of rows between the initial position row and the final position row.");

            #endregion
            #endregion

            #region Call NspiUpdateStat with Delta field of STAT block set to 0 to get the first row position.
            stat.InitiateStat();
            stat.Delta = 0;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            // Record the first row position of the table.
            uint firstRowPosition = stat.CurrentRec;
            #endregion

            #region Call NspiUpdateStat with Delta field of STAT block set to a negative value whose absolute value is greater than TotalRecs to move Current Position to be before the first row of the table.
            stat.Delta = -((int)stat.TotalRecs + 1);
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1744");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1744
            // In this step, server moves the current position to be before the first row of the table. MS-OXNSPI_R1744 can be verified if stat.CurrentRec equals the first row position gotten from the previous step.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                firstRowPosition,
                stat.CurrentRec,
                1744,
                @"[In Absolute Positioning] [If applying the Delta as described in step 6 would move the Current Position to be before the first row of the table, the server] sets the CurrentRec to the Minimal Entry ID of the object that occupies the first row of the table.");
            #endregion
            #endregion

            #region Call NspiUpdateStat with CurrentRec set to MID_CURRENT and Delta field of STAT block to the value greater than TotalRecs to move Current Position to be after the end the table.
            stat.CurrentRec = (uint)MinimalEntryID.MID_CURRENT;
            stat.Delta = (int)stat.TotalRecs + 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R601");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R601
            // Now the current position is moved to be the end of the table, so MS-OXNSPI_R601 can be verified if NumPos equals TotalRecs.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                stat.TotalRecs,
                stat.NumPos,
                601,
                @"[In Fractional Positioning] The server reports this [Numeric Position of the Current Position of the table] in the NumPos field of the STAT structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1748");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1748
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_END_OF_TABLE,
                stat.CurrentRec,
                1748,
                @"[In Fractional Positioning] [If applying Delta as described in step 8 would move the Current Position to be after the end of the table, the server] sets the CurrentRec to the value MID_END_OF_TABLE.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R81");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R81
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_END_OF_TABLE,
                stat.CurrentRec,
                81,
                @"[In Positioning Minimal Entry IDs] MID_END_OF_TABLE (0x00000002): Specifies the position after the last row in the current address book container.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R768");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R768
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_END_OF_TABLE,
                stat.CurrentRec,
                768,
                @"[In NspiUpdateStat] The NspiUpdateStat method updates the STAT block that represents position in a table to reflect positioning changes requested by the client.");

            // pStat is defined to the type of STAT in this test suite, and the CurrentRec describes the logical position. So this requirement can be captured if code can reach here
            this.Site.CaptureRequirement(
                773,
                @"[In NspiUpdateStat] pStat: A pointer to a STAT block describing a logical position in a specific address book container.");
            #endregion
            #endregion

            #region Call NspiUpdateStat with CurrentRec set to MID_CURRENT and Delta field of STAT block to the value whose absolute value is greater than TotalRecs to move Current Position to be before the first row of the table.
            stat.CurrentRec = (uint)MinimalEntryID.MID_CURRENT;
            stat.Delta = -(int)stat.TotalRecs - 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1746");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1746
            // Now the position is before the first row in the current address book container since stat.Delta is set to -TotalRecs - 1.
            // So MS-OXNSPI_R1746 can be verified if stat.CurrentRec == firstRowPosition.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                firstRowPosition,
                stat.CurrentRec,
                1746,
                @"[In Fractional Positioning] [If applying the Delta as described in step 8 would move the Current Position to be before the beginning of the table, the server] sets the CurrentRec field to the Minimal Entry ID of the object occupying the first row of the table.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to Absolute Positioning when updating STAT block.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC04_UpdateStatAbsolutePosition()
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

            #region Call NspiUpdateStat with the CurrentRec field set to MID_BEGINNING_OF_TABLE (0x00) rather than MID_CURRENT (0x1) to point to the absolute position, and Delta set to 1 to move forward one row towards the end of the table.
            // Record the position before updating stat in this step.
            uint previousPosition = stat.NumPos;
            stat.InitiateStat();
            stat.Delta = 1;
            uint reserved = 0;
            int? delta = 0;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R549. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R549
            // If the currentPosition greater than the position before moving, it specifies that the current position is moved toward the end of the table.
            this.Site.CaptureRequirementIfIsTrue(
                stat.NumPos > previousPosition,
                549,
                @"[In Absolute Positioning] [step 6] If the value of Delta is positive, the Current Position is moved toward the end of the table.");
            #endregion

            uint previousTotalRecs = stat.TotalRecs;
            previousPosition = stat.NumPos;
            stat.InitiateStat();
            stat.NumPos = 1;
            stat.Delta = 1;
            delta = 0;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R326001");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R326001
            // NumPos returned are same regardless of the value of NumPos set in request, this requirement can be captured.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                previousPosition,
                stat.NumPos,
                326001,
                @"[In STAT] [NumPos] If absolute positioning, as specified in section 3.1.4.5.1, is used, the value of this field specified by the client will be ignored by the server. ");
            #endregion

            stat.InitiateStat();
            stat.TotalRecs = 1;
            stat.Delta = 1;
            delta = 0;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R331001");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R331001
            this.Site.CaptureRequirementIfAreEqual<uint>(
                previousTotalRecs,
                stat.TotalRecs,
                331001,
                @"[In STAT] [TotalRecs] If absolute positioning, as specified in section 3.1.4.5.1, is used, the value of this field specified by the client will be ignored by the server. ");
            #endregion

            #endregion

            #region Call NspiUpdateStat with Delta set to 0 to keep the position in the table.
            // Record the position before updating stat in this step.
            previousPosition = stat.NumPos;
            stat.Delta = 0;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R550. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R550
            // If the currentPosition equals the position before moving, it specifies that the current position is not changed.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                previousPosition,
                stat.NumPos,
                550,
                @"[In Absolute Positioning] [step 6] A Delta with the value 0 results in no change to the Current Position.");

            #endregion
            #endregion

            #region Call NspiUpdateStat with Delta set to -1 to move forward one row towards the beginning of the table.
            // Record the position before updating stat in this step.
            previousPosition = stat.NumPos;
            stat.Delta = -1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R548. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R548
            // If the currentPosition less than the position before moving, it specifies that the current position is moved toward the beginning of the table.
            this.Site.CaptureRequirementIfIsTrue(
                stat.NumPos < previousPosition,
                548,
                @"[In Absolute Positioning] [step 6] If the value of Delta is negative, the Current Position is moved toward the beginning of the table.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R791.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R791
            // MS-OXNSPI_R548 has been verified with the CurrentRec field set to MID_BEGINNING_OF_TABLE and the server locates that row as the initial position row, MS-OXNSPI_R791 can be verified directly.
            this.Site.CaptureRequirement(
                791,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 5: The server locates the initial position row in the table specified by the ContainerID field of the input parameter pStat as follows:] If the row specified by the CurrentRec field of the input parameter pStat is not MID_CURRENT, the server locates that row as the initial position row using the absolute position, as specified in section 3.1.4.5.1. ");

            #endregion
            #endregion

            #region Call NspiUpdateStat with Delta set to -1 to move the current position to be before the first row of the table.
            stat.Delta = -1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1743. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1743
            // In this step, the current position should be moved to be before the first row of the table.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE,
                stat.NumPos,
                1743,
                @"[In Absolute Positioning] [step 7] If applying the Delta as described in step 6 would move the Current Position to be before the first row of the table, the server sets the Current Position to the first row of the table.");
            #endregion
            #endregion

            #region Call NspiUpdateStat set Delta to the value greater than TotalRecs to move the current position to be after the end of the table.
            stat.Delta = (int)stat.TotalRecs + 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R554. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R554
            // In this step, the current position is moved to be after the end of the table.
            // If the currentPosition equals the total number of rows, it indicates that server sets the current position to a location which is one row past the last valid row of the table.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                stat.TotalRecs,
                stat.NumPos,
                554,
                @"[In Absolute Positioning] [step 8] If applying the Delta as described in step 6 would move the Current Position to be after the end of the table, the server sets the Current Position to a location one row past the last valid row of the table.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R556");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R556
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_END_OF_TABLE,
                stat.CurrentRec,
                556,
                @"[In Absolute Positioning] [step 9] The server sets the field CurrentRec to the Minimal Entry ID of the object occupying the row specified by the Current Position.");

            if (Common.IsRequirementEnabled(1992, this.Site))
            {
                // All values got from NumPos field are accurate values rather than approximate values, so MS-OXNSPI_R1992 can be verified directly.
                this.Site.CaptureRequirement(
                    1992,
                    @"[In Appendix B: Product Behavior] Implementation doesn't report an approximate value for the Numeric Position. (Exchange server 2010 and above follow this behavior.)");
            }

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to Fractional Positioning when updating STAT block.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC05_UpdateStatFractionalPosition()
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

            #region Call NspiUpdateStat with the CurrentRec field set to MID_CURRENT to point to the fractional position and Delta set to 1 to move forward one row towards the end of the table.
            stat.InitiateStat();
            stat.Delta = 1;
            stat.CurrentRec = (uint)MinimalEntryID.MID_CURRENT;
            uint reserved = 0;
            int? delta = 0;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R588. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R588
            // If the currentPosition greater than the first position, it specifies that server moves towards the end of the table.
            this.Site.CaptureRequirementIfIsTrue(
                stat.NumPos > 0,
                588,
                @"[In Fractional Positioning] If the value of Delta is positive, the Current Position is moved toward the end of the table.");

            #endregion
            #endregion

            #region Call NspiUpdateStat with Delta set to 0 to keep the position in the table.
            // Record the position before updating stat in this step.
            uint previousPosition = stat.NumPos;
            stat.Delta = 0;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R589. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R589
            // If the current position equals the previous position, it specifies that server keeps the position in the table.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                previousPosition,
                stat.NumPos,
                589,
                @"[In Fractional Positioning] A Delta field with the value 0 results in no change to the Current Position.");
            #endregion
            #endregion

            #region Call NspiUpdateStat with Delta set to -1 to move forward one row towards the beginning of the table.
            int supposedValueOfTotalRecs = (int)stat.TotalRecs;
            stat.Delta = -1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R328");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R328
            // In this step, the current position is moved to the initial position of the table. So if stat.NumPos is 0, this requirement can be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                stat.NumPos,
                328,
                @"[In STAT] [NumPos] This value [the approximate fractional position] is a zero index; the first element in a table has the numeric position 0.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R598");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R598
            // In this step, the current position is moved to the initial position of the table. So if stat.NumPos is 0, this requirement can be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                stat.NumPos,
                598,
                @"[In Fractional Positioning] This numeric row [the numeric row of the Current Position] is 0-based.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R587. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R587
            // In the previous steps, the current position is moved one row from the beginning of the table towards the end of the table. Delta is set to -1 in this step, so if currentPosition equals 0, 
            // it indicates server moves towards the beginning of the table, and MS-OXNSPI_R587 can be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                stat.NumPos,
                587,
                @"[In Fractional Positioning] If the value of Delta is negative, the Current Position is moved toward the beginning of the table.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R793");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R793
            // MS-OXNSPI_R587, MS-OXNSPI_588 and MS-OXNSPI_R589 have been verified, MS-OXNSPI_R793 can be verified directly.
            this.Site.CaptureRequirement(
                793,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 5: The server locates the initial position row in the table specified by the ContainerID field of the input parameter pStat as follows:] If the row specified by the CurrentRec field of the input parameter pStat is MID_CURRENT, the server locates the initial position row using the fractional position specified in the NumPos field of the input parameter pStat as specified in section 3.1.4.5.2.");

            // Calculate the intended numeric position.
            float intendedNumericPosition = stat.TotalRecs * stat.NumPos / supposedValueOfTotalRecs;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R577. The intended numeric position is {0}, the actual NumPos is {1}.", intendedNumericPosition, stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R577
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)intendedNumericPosition,
                stat.NumPos,
                577,
                @"[In Fractional Positioning] [step 4] The server calculates the Intended Numeric Position in the table as the TotalRecs reported by the server multiplied by the NumPos field of the STAT structure divided by the value of TotalRecs as specified by the client.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R578. The intended numeric position is {0}, the actual NumPos is {1}.", intendedNumericPosition, stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R578
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)intendedNumericPosition,
                stat.NumPos,
                578,
                @"[In Fractional Positioning] [step 4] The value is truncated to its integral part.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R584");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R584
            // In step 6 of Fractional Positioning: The server identifies the numeric row of the Current Position in the sorted table.
            // In this step, the first row is chosen to be moved to, so if stat.NumPos points to the first row in the table, MS-OXNSPI_R584 can be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE,
                stat.NumPos,
                584,
                @"[In Fractional Positioning] The server moves the Current Position to the row chosen in step 6.");
            #endregion
            #endregion

            #region Call NspiUpdateStat with Delta set to -1 to move the Current Position to be before the beginning of the table.
            stat.Delta = -1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1745. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1745
            // In this step, the current position should be moved to be before the first row of the table.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE,
                stat.NumPos,
                1745,
                @"[In Fractional Positioning] If applying the Delta as described in step 8 would move the Current Position to be before the beginning of the table, the server sets the Current Position to the beginning of the table.");
            #endregion
            #endregion

            #region Call NspiUpdateStat with Delta set to TotalRecs + 1 to move the Current Position to be after the end of the table.
            supposedValueOfTotalRecs = (int)stat.TotalRecs - 1; // A nonzero value which is less than the actual value of stat.TotalRecs used to build the pre-condition for capture code R580.
            stat.Delta = (int)stat.TotalRecs + 1;
            uint totalRecsSpecifiedByClient = 10;
            stat.TotalRecs = totalRecsSpecifiedByClient;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R797", "The value of NumPos is {0}, the value of TotalRecs is {1}.", stat.NumPos, stat.TotalRecs);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R797
            // According to section 3.1.4.5.2 of this Open Specification, the NumPos field is calculated as follows: 
            // TotalRecs reported by the server multiplied by the NumPos field of the STAT structure divided by the value of TotalRecs
            // as specified by the client. And then, approximate fractional position is the integral part.
            this.Site.CaptureRequirementIfIsTrue(
                stat.NumPos == stat.NumPos * stat.TotalRecs / totalRecsSpecifiedByClient || stat.NumPos == stat.TotalRecs,
                797,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 9] The server MUST set the NumPos field of the parameter pStat to the approximate numeric position of the current row of the current address book container according to section 3.1.4.5.2.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R327
            // MS-OXNSPI_R797 has been verified, MS-OXNSPI_R327 can be verified directly.
            this.Site.CaptureRequirement(
                327,
                @"[In STAT] [NumPos] The server sets this field to specify the approximate fractional position at the end of an NSPI method.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1747. The position after moving is {0}.", stat.NumPos);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1747
            // If the currentPosition equals the TotalRecs+1, it specifies that server sets the Current Position to a location one row past the last valid row of the table.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                stat.TotalRecs,
                stat.NumPos,
                1747,
                @"[In Fractional Positioning] If applying Delta as described in step 8 would move the Current Position to be after the end of the table, the server sets the Current Position to a location one row past the last valid row of the table.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R595");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R595
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_END_OF_TABLE,
                stat.CurrentRec,
                595,
                @"[In Fractional Positioning] The server sets the field CurrentRec to the Minimal Entry ID of the object occupying the row specified by the Current Position.");

            // Calculate the intended numeric position.
            intendedNumericPosition = stat.TotalRecs * stat.NumPos / supposedValueOfTotalRecs;

            if (intendedNumericPosition > stat.TotalRecs)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R580");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R580
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    stat.TotalRecs,
                    stat.NumPos,
                    580,
                    @"[In Fractional Positioning] [step 5] If the Intended Numeric Position thus calculated is greater than TotalRecs, the intended Intended Numeric Position is TotalRecs (that is, the last row in the table).");
            }
            else
            {
                Site.Assert.Fail("The Intended Numeric Position {0} not as expected to be greater than TotalRecs {1}.", intendedNumericPosition, stat.TotalRecs);
            }

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetProps operation returning success with different flags.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC06_GetPropsSuccessWithDifferentFlags()
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
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiUpdateStat to update the STAT block to make CurrentRec point to the first row of the table.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiGetProps method with dwFlags set to fEphID.
            PropertyTagArray_r prop = new PropertyTagArray_r
            {
                CValues = 2
            };
            prop.AulPropTag = new uint[prop.CValues];
            prop.AulPropTag[0] = (uint)AulProp.PidTagEntryId;
            prop.AulPropTag[1] = (uint)AulProp.PidTagDisplayName;
            PropertyTagArray_r? propTags = prop;

            uint flagsOfGetProps = (uint)RetrievePropertyFlag.fEphID;
            PropertyRow_r? rows;

            // Save the CurrentRec value of stat structure
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetProps should return Success.");
            Site.Assert.IsNotNull(rows, "Rows should not be null. The row number is {0}.", rows == null ? 0 : rows.Value.CValues);

            // If input parameter dwFlags contains the bit flag fEphID and the PidTagEntryId property is present in the list of proptags, 
            // the server MUST return the values of the PidTagEntryId property in the Ephemeral Entry ID format, as specified in section 2.3.8.2.
            EphemeralEntryID entryID1 = AdapterHelper.ParseEphemeralEntryIDFromBytes(rows.Value.LpProps[0].Value.Bin.Lpb);

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R105");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R105
            // The ID Type of EphemeralEntryID is 0x87.
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x87,
                entryID1.IDType,
                105,
                @"[In Retrieve Property Flags] fEphID (0x00000002): Client requires that the server MUST return Entry ID values in Ephemeral Entry ID form.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R905");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R905
            Site.CaptureRequirementIfAreNotEqual<uint>(
                0x0,
                rows.Value.CValues,
                905,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 10] If the server can locate the object specified in the CurrentRec field of the input parameter pStat, the server MUST return values associated with this object [the object specified in the CurrentRec field of the input parameter pStat].");

            #endregion
            #endregion

            #region Call NspiGetProps method with dwFlags set to fSkipObjects.
            flagsOfGetProps = (uint)RetrievePropertyFlag.fSkipObjects;

            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetProps method should succeed.");
            PermanentEntryID entryID = AdapterHelper.ParsePermanentEntryIDFromBytes(rows.Value.LpProps[0].Value.Bin.Lpb);

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R901");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R901
            // The entryID is parsed as Permanent Entry ID format, if it parsed successfully, then the PidTagEntryId property is in the specified format.
            Site.CaptureRequirementIfIsNotNull(
                entryID,
                901,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] If input parameter dwFlags does not contain the bit flag fEphID and the PidTagEntryId property is present in the list of proptags, the server MUST return the values of the PidTagEntryId property in the Permanent Entry ID format, as specified in section 2.2.9.3.");

            this.VerifyServerUsesListSpecifiedBypPropTagsInNspiGetProps(propTags, rows);

            #endregion
            #endregion

            #region Call NspiGetProps method with dwFlags set to fEphID again to check EphemeralEntryID.
            flagsOfGetProps = (uint)RetrievePropertyFlag.fEphID;

            // Get the specific CurrentRec.
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetProps method should succeed.");
            EphemeralEntryID entryID2 = AdapterHelper.ParseEphemeralEntryIDFromBytes(rows.Value.LpProps[0].Value.Bin.Lpb);

            #region Capture code
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1994. The DisplayType of entryID1 is {0}. The Mid of entryID1 is {1}. The DisplayType of entryID2 is {2}. The Mid of entryID2 is {3}.",
                entryID1.DisplayType,
                entryID1.Mid,
                entryID2.DisplayType,
                entryID2.Mid);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1994
            Site.CaptureRequirementIfIsTrue(
                entryID1.Compare(entryID2),
                1994,
                @"[In EphemeralEntryID] A server MUST NOT change an object's Ephemeral Entry ID during two calls.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1995
            // MS-OXNSPI_R1994 has been verified, so MS-OXNSPI_R1995 can be verified directly.
            Site.CaptureRequirement(
                1995,
                @"[In Object Identity] A server MUST NOT change an object's Ephemeral Identifier during two calls.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R605
            // MS-OXNSPI_R1994 has been verified, MS-OXNSPI_R605 can be verified directly.
            Site.CaptureRequirement(
                605,
                "[In Object Identity] Ephemeral Identifier: Specifies a specific object in a single NSPI session.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1706
            // If the codes can reach here, the Ephemeral Identifier of a object must be the same, which is verified in R1994.
            // Because the Minimal Identifier is included in the Ephemeral Identifier, so this requirement can be captured directly.
            Site.CaptureRequirement(
                1706,
                @"[In Object Identity] A server MUST NOT change an object's Minimal Entry ID between two calls in the lifetime of an NSPI session.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R899");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R899
            // The entryID is parsed as Ephemeral Entry ID format, if it parsed successfully, then the PidTagEntryId property is in the specified format.
            Site.CaptureRequirementIfIsNotNull(
                entryID2,
                899,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] If input parameter dwFlags contains the bit flag fEphID and the PidTagEntryId property is present in the list of proptags, the server MUST return the values of the PidTagEntryId property in the Ephemeral Entry ID format, as specified in section 2.2.9.2.");
            #endregion
            #endregion

            #region Call NspiGetProps method with dwFlags set to one value that is not fEphID (0x2) or fSkipObjects (0x1).
            STAT savedStat = stat;
            flagsOfGetProps = 4;
            PropertyRow_r? rows1;
            ErrorCodeValue result1 = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result1, "NspiGetProps method should succeed.");
            #endregion

            #region Call NspiGetProps method with dwFlags set to another value that is not fEphID (0x2) or fSkipObjects (0x1).
            flagsOfGetProps = 8;
            PropertyRow_r? rows2;
            ErrorCodeValue result2 = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, savedStat, propTags, out rows2);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result2, "NspiGetProps method should succeed.");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1925");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1925
            // Check whether the returned rows are equal.
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowEqual(rows1, rows2),
                1925,
                @"[In NspiGetProps] If dwFlags is set to different values other than the bit flags fEphID and fSkipObjects, server will return the same result.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetPropList method using parameters of NspiGetProps method.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC07_GetPropsAndNspiGetPropList()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();
            if (this.Transport == "mapi_http")
            {
                Site.Assert.Inconclusive("This case can not run for that the NspiGetPros method cannot return all the properties same as the NspiGetPropList method using mapi_http transport.");
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

            #region Call NspiUpdateStat to update the STAT block to make CurrentRec point to the first row of the table.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiGetProps method with proptags set to null.
            PropertyTagArray_r? propTags = null;
            uint flagsOfGetProps = (uint)RetrievePropertyFlag.fEphID;
            PropertyRow_r? rows;
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R915");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R915
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                915,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 13] If no other return values have been specified by these constraints [constraints 1-12], the server MUST return the return value ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R859");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R859
            this.Site.CaptureRequirementIfIsNotNull(
                rows,
                859,
                @"[In NspiGetProps] The NspiGetProps method returns an address book row that contains a set of the properties and values that exist on an object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R868");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R868
            // Method NspiGetProps is getting properties from the address book container, so if the rows returned are not null, it specifies that ppRows contains the address book container row.
            this.Site.CaptureRequirementIfIsNotNull(
                rows,
                868,
                @"[In NspiGetProps] [ppRows] Contains the address book container row the server returns in response to the request.");
            #endregion
            #endregion

            #region Call NspiGetPropList method using the input parameters the same as those of NspiGetProps.
            PropertyTagArray_r? propTagsOfGetPropList;
            this.Result = this.ProtocolAdatper.NspiGetPropList(flagsOfGetProps, stat.CurrentRec, stat.CodePage, out propTagsOfGetPropList);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetPropList should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1693");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1693
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsRowSubjectToPropTags(propTagsOfGetPropList, rows),
                1693,
                @"[In NspiGetProps] [constraint 5] The server MUST construct a proptag list that is exactly the same list that would be returned to the client in the pPropTags output parameter of the NspiGetPropList method, as specified in section 3.1.4.1.6, using the following parameters as inputs to the NspiGetPropList method: 
                The NspiGetProps parameter hRpc is used as the NspiGetPropList parameter hRpc. 
                The NspiGetProps parameter dwFlags is used as the NspiGetPropList parameter dwFlags. 
                The CurrentRec field of the NspiGetProps parameter pStat is used as the NspiGetPropList parameter dwMId.
                The CodePage field of the NspiGetProps parameter pStat is used as the NspiGetPropList parameter CodePage.");
            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetProps operation with dwFlags containing fEphID.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC08_GetPropsSuccessWithFlagsContainsfEphID()
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
            // The expected CurrentRec value for administrator.
            uint expectedRec = 0;
            stat.InitiateStat();
            uint reserved = 0;
            int? delta = 0;
            stat.Delta = 0;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            uint[] recs = new uint[stat.TotalRecs];
            recs[0] = stat.CurrentRec;

            // Get each CurrentRec values.
            for (int i = 1; i < recs.Length; i++)
            {
                stat.Delta = 1;
                this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
                recs[i] = stat.CurrentRec;
            }
            #endregion

            #region Call NspiGetProps with input parameter proptags not set to null.
            for (int i = 0; i < recs.Length; i++)
            {
                PropertyRow_r? rows;
                uint flag = (uint)RetrievePropertyFlag.fEphID;
                PropertyTagArray_r prop = new PropertyTagArray_r
                {
                    CValues = 2
                };
                prop.AulPropTag = new uint[prop.CValues];
                prop.AulPropTag[0] = (uint)AulProp.PidTagEntryId;
                prop.AulPropTag[1] = (uint)AulProp.PidTagDisplayName;
                PropertyTagArray_r? propTags = prop;

                stat.InitiateStat();
                stat.CurrentRec = recs[i];
                this.Result = this.ProtocolAdatper.NspiGetProps(flag, stat, propTags, out rows);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetProps should return Success.");

                // Save display name and corresponding CurrentRec value in hash table.
                string name = System.Text.Encoding.UTF8.GetString(rows.Value.LpProps[1].Value.LpszA);
                string administratorName = Common.GetConfigurationPropertyValue("User1Name", this.Site);

                // In environment, the user administrator is added x509 certificate. So get CurrentRec value of administrator to get PidTagAddressBookX509Certificate type.
                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(administratorName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    expectedRec = recs[i];
                    break;
                }
            }

            PropertyTagArray_r propInstance = new PropertyTagArray_r
            {
                CValues = 2
            };
            propInstance.AulPropTag = new uint[propInstance.CValues];
            propInstance.AulPropTag[0] = (uint)AulProp.PidTagEntryId;
            propInstance.AulPropTag[1] = (uint)AulProp.PidTagAddressBookX509Certificate;
            PropertyTagArray_r? propTagsOfGetProps = propInstance;

            uint flagsOfGetProps = (uint)RetrievePropertyFlag.fEphID;
            PropertyRow_r? rowsOfGetProps;

            stat.CurrentRec = expectedRec;
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTagsOfGetProps, out rowsOfGetProps);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetProps method should succeed.");

            #region Capture

            this.VerifyServerUsesListSpecifiedBypPropTagsInNspiGetProps(propTagsOfGetProps, rowsOfGetProps);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R911");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R911
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsRowSubjectToPropTags(propTagsOfGetProps, rowsOfGetProps),
                911,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 12] Subject to the prior constraints, the server constructs a list of properties and their values as a single PropertyRow_r structure with a one-to-one order preserving correspondence between the values in the proptag list specified by input parameters and the returned properties and values in the RowSet.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R913");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R913
            // The rows is used as the instance parameter of ppRows, if it is not null, it illustrates that the server returned the structure.
            Site.CaptureRequirementIfIsNotNull(
                rowsOfGetProps,
                913,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 12] The server MUST return this RowSet in the output parameter ppRows.");

            #endregion Capture

            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetProps operation requests with duplicate properties.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC09_GetPropsSuccessWithDuplicateProperties()
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

            #region Call NspiGetProps with the input parameter propTags that has duplicate properties.
            PropertyTagArray_r prop = new PropertyTagArray_r
            {
                CValues = 2
            };
            prop.AulPropTag = new uint[prop.CValues];
            prop.AulPropTag[0] = (uint)AulProp.PidTagDisplayType;
            prop.AulPropTag[1] = (uint)AulProp.PidTagDisplayType;
            PropertyTagArray_r? propTags = prop;

            uint flagsOfGetProps = (uint)RetrievePropertyFlag.fEphID;
            PropertyRow_r? rows;

            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetProps method should succeed.");

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R912");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R912
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyValueEqual(rows.Value.LpProps[0], rows.Value.LpProps[1]),
                912,
                @"[In NspiGetProps] [Server Processing Rules: Upon receiving message NspiGetProps, the server MUST process the data from the message subject to the following constraints:] [Constraint 12] If there are duplicate properties in the proptag list, the server MUST create duplicate values in the parameter RowSet.");

            this.VerifyServerUsesListSpecifiedBypPropTagsInNspiGetProps(propTags, rows);
            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to comparing NspiSeekEntries operation and NspiQueryRows operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC10_SeekEntriesSuccessCompareWithNspiQueryRows()
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

            #region Call NspiGetPropList method with dwFlags set to fEphID and CodePage field of STAT block not set to CP_WINUNICODE.
            uint flagsOfGetPropList = (uint)RetrievePropertyFlag.fEphID;
            PropertyTagArray_r? propTagsOfGetPropList;
            uint codePage = (uint)RequiredCodePage.CP_TELETEX;

            this.Result = this.ProtocolAdatper.NspiGetPropList(flagsOfGetPropList, stat.CurrentRec, codePage, out propTagsOfGetPropList);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetPropList should return Success!");
            #endregion

            #region Call NspiSeekEntries method with requesting one property tag.
            uint reservedOfSeekEntries = 0;
            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = (uint)0x00
            };
            string displayName;
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                displayName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
            }
            else
            {
                displayName = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "\0";
            }

            target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);

            PropertyTagArray_r tags = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagDisplayName
                }
            };
            PropertyTagArray_r? propTagsOfSeekEntries = tags;

            PropertyTagArray_r? tableOfSeekEntries = null;
            PropertyRowSet_r? rowsOfSeekEntries;

            uint beforePosition = stat.CurrentRec;

            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiSeekEntries method should succeed.");

            #region Capture

            bool isTheValueGreaterThanOrEqualTo = true;
            for (int i = 0; i < rowsOfSeekEntries.Value.CRows; i++)
            {
                // Check whether the returned property value is greater than or equal to the input property PidTagDisplayName.
                string dispalyNameReturnedFromServer = System.Text.Encoding.UTF8.GetString(rowsOfSeekEntries.Value.ARow[i].LpProps[0].Value.LpszA);

                displayName = displayName.TrimEnd('\0');

                // Value Condition greater than zero indicates the returned property value follows the input PidTagDisplayName property value.
                int result = string.Compare(dispalyNameReturnedFromServer, displayName, StringComparison.Ordinal);                
                if (result < 0)
                {
                    this.Site.Log.Add(LogEntryKind.Debug, "The display name returned from server is {0}, the expected display name is {1}", dispalyNameReturnedFromServer, displayName);

                    isTheValueGreaterThanOrEqualTo = false;
                    break;
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R998");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R998
            this.Site.CaptureRequirementIfIsTrue(
                isTheValueGreaterThanOrEqualTo,
                998,
                @"[In NspiSeekEntries] The NspiSeekEntries method searches for and sets the logical position in a specific table to the first entry greater than or equal to a specified value.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1050. The position before calling method NspiSeekEntries is {0}, the position after calling method NspiSeekEntries is {1}.", beforePosition, stat.CurrentRec);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1050
            this.Site.CaptureRequirementIfIsTrue(
                rowsOfSeekEntries != null && stat.CurrentRec != beforePosition,
                1050,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] If a qualifying row was found, the server MUST update the position information in the parameter pStat.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1076");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1076
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1076,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 16] If no other return values have been specified by these constraints [constraints 1-15], the server MUST return the return value ""Success"".");

            #endregion Capture
            #endregion

            #region Call NspiQueryRows method using parameters of NspiSeekEntries method.
            uint flagsOfQueryRows1 = (uint)RetrievePropertyFlag.fEphID;
            uint tableCount1 = 0;
            uint[] table1 = null;
            uint count1 = rowsOfSeekEntries.Value.CRows;
            PropertyRowSet_r? rowsOfQueryRows1;
            PropertyTagArray_r? propTags1 = tags;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows1, ref stat, tableCount1, table1, count1, propTags1, out rowsOfQueryRows1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1695");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1695
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowSetEqual(rowsOfSeekEntries, rowsOfQueryRows1),
                1695,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [constraint 15] This PropertyRowSet_r MUST be exactly the same PropertyRowSet_r that would be returned in the ppRows parameter of a call to the NspiQueryRows method with the following parameters:
                The NspiSeekEntries parameter hRpc is used as the NspiQueryRows parameter hRpc.
                The value fEphID is used as the NspiQueryRows parameter dwFlags.
                The NspiSeekEntries output parameter pStat (as modified by the prior constraints) is used as the NspiQueryRows parameter pStat.
                If the NspiSeekEntries input parameter lpETable is NULL, the value 0 is used as the NspiQueryRows parameter dwETableCount, and the value NULL is used as the NspiQueryRows parameter lpETable.
                If the NspiSeekEntries input parameter lpETable is NULL, the server MUST choose a value for the NspiQueryRows parameter Count. The Exchange Server NSPI Protocol does not prescribe any particular algorithm. The server MUST use a value greater than 0.
                The NspiSeekEntries parameter pPropTags is used as the NspiQueryRows parameter pPropTags.");
            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiSeekEntries operation returning success with null table.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC11_SeekEntriesSuccessWithNullTable()
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

            #region Call NspiSeekentries method with input parameter Reserved set to 0 and input parameter lpETable set to null.
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

            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiSeekEntries method should succeed.");

            #region Capture

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1055: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
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

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1055
            bool isVerifyR1055 = inputStat.CodePage == stat.CodePage
                                && inputStat.ContainerID == stat.ContainerID
                                && inputStat.Delta == stat.Delta
                                && inputStat.SortLocale == stat.SortLocale
                                && inputStat.SortType == stat.SortType
                                && inputStat.TemplateLocale == stat.TemplateLocale;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1055,
                1055,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] The server MUST NOT modify any other fields [CodePage, ContainerID, Delta, SortLocale, SortType, TemplateLocale] of the parameter pStat.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1051, the CurrentRec of stat is {0}", stat.CurrentRec);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1051
            // According to section 2.3.8.1 of Open Specification MS-OXNSPI, Minimal Entry IDs with values less than 0x00000010 are used by clients as signals to 
            // trigger specific behaviors in specific NSPI methods. So if the stat.CurrentRec >= 0x10, it indicates that the Minimal Entry ID is returned by server, and then MS-OXNSPI_R1051 can be verified.
            Site.CaptureRequirementIfIsTrue(
                stat.CurrentRec >= 0x10,
                1051,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] [If a qualifying row was found,] The server MUST set CurrentRec field of the parameter pStat to the Minimal Entry ID of the qualifying row.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1054");

            uint objectNum = this.SutControlAdapter.GetNumberOfAddressBookObject();

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1054
            Site.CaptureRequirementIfAreEqual<uint>(
                objectNum,
                stat.TotalRecs,
                1054,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] The TotalRecs field of the parameter pStat MUST be set to the accurate number of records in the table used.");

            #endregion
            #endregion

            #region Call NspiSeekentries method again with input parameter Reserved set to 1.
            PropertyRowSet_r? rowsOfSeekEntries1;
            reservedOfSeekEntries = 0x1;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref inputStat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiSeekEntries should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1702");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1702
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowSetEqual(rowsOfSeekEntries, rowsOfSeekEntries1),
                1702,
                @"[In NspiSeekEntries] If this field[Reserved] is set to different values, the server will return the same result.");
            #endregion Capture

            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiQueryRows operation returning success.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC12_QueryRowsSuccess()
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

            #region Call NspiQueryRows with null table.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            PropertyRowSet_r? rowsOfQueryRows;

            // Set the input parameter lpETable to NULL. 
            uint tableCount = 0;
            uint[] table = null;
            uint count = 1; // Set the expected row count to be returned.
            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 16,
                AulPropTag = new uint[16]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType,
                    (uint)AulProp.PidTagInitialDetailsPane,
                    (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                    (uint)AulProp.PidTagAddressBookContainerId,
                    (uint)AulProp.PidTagInstanceKey,
                    (uint)AulProp.PidTagSearchKey,
                    (uint)AulProp.PidTagRecordKey,
                    (uint)AulProp.PidTagAddressType,
                    (uint)AulProp.PidTagEmailAddress,
                    (uint)AulProp.PidTagTemplateid,
                    (uint)AulProp.PidTagTransmittableDisplayName,
                    (uint)AulProp.PidTagMappingSignature,
                    (uint)AulProp.PidTagAddressBookObjectDistinguishedName
                }
            };
            PropertyTagArray_r? propTags = propTagsInstance;

            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");
            Site.Assert.IsNotNull(rowsOfQueryRows.Value.ARow, "The returned rows should not be empty. The row number is {0}.", rowsOfQueryRows == null ? 0 : rowsOfQueryRows.Value.CRows);

            #region Capture

            // Add the debug information 
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R927");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R927 
            this.Site.CaptureRequirementIfAreEqual<uint>(
                count,
                rowsOfQueryRows.Value.CRows,
                927,
                @"[In NspiQueryRows] Count: A DWORD value that contains the number of rows the client is requesting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R916");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R916
            // The returned rows are specified in rowsOfQueryRows.Value.ARow. So if rowsOfQueryRows.Value.ARow is not null, it specifies that method NspiQueryRows
            // returns to the client a number of rows.
            this.Site.CaptureRequirementIfIsNotNull(
                rowsOfQueryRows.Value.ARow,
                916,
                @"[In NspiQueryRows] The NspiQueryRows method returns to the client a number of rows from a specified table.");

            // pStat is defined to the type of STAT in this test suite, and the CurrentRec describes the logical position. So this requirement can be captured if code can reach here.
            this.Site.CaptureRequirement(
                922,
                @"[In NspiQueryRows] pStat: A pointer to a STAT block that describes a logical position in a specific address book container.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R955: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
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

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R955
            // According to the description in section 3.1.4.1.8, if the server is not using the table specified by the input parameter pStat, the server must not update the status of the table. 
            // So if the returned STAT is updated, it is believed that the server uses the table specified by the input parameter pStat.
            bool isVerifyR955 = inputStat.CodePage != stat.CodePage
                                || inputStat.ContainerID != stat.ContainerID
                                || inputStat.CurrentRec != stat.CurrentRec
                                || inputStat.Delta != stat.Delta
                                || inputStat.NumPos != stat.NumPos
                                || inputStat.SortLocale != stat.SortLocale
                                || inputStat.SortType != stat.SortType
                                || inputStat.TemplateLocale != stat.TemplateLocale
                                || inputStat.TotalRecs != stat.TotalRecs;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR955,
                955,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] If the input parameter lpETable is NULL, the server MUST use the table specified by the input parameter pStat when constructing the return parameter ppRows.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R988: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
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

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R988
            // If the output status of the table is not equal to the input status, it means that server has updated the status of the table.
            bool isVerifyR988 = inputStat.CodePage != stat.CodePage
                                || inputStat.ContainerID != stat.ContainerID
                                || inputStat.CurrentRec != stat.CurrentRec
                                || inputStat.Delta != stat.Delta
                                || inputStat.NumPos != stat.NumPos
                                || inputStat.SortLocale != stat.SortLocale
                                || inputStat.SortType != stat.SortType
                                || inputStat.TemplateLocale != stat.TemplateLocale
                                || inputStat.TotalRecs != stat.TotalRecs;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR988,
                988,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 16] If the server is using the table specified by the input parameter pStat, the server MUST update the status of the table.");

            // Whether GUID maintain or not
            bool isGUIDMaintain = true;

            foreach (PropertyRow_r row in rowsOfQueryRows.Value.ARow)
            {
                EphemeralEntryID entryID = AdapterHelper.ParseEphemeralEntryIDFromBytes(row.LpProps[0].Value.Bin.Lpb);
                if (!AdapterHelper.AreTwoByteArrayEqual(serverGuid.Value.Ab, entryID.ProviderUID.Ab))
                {
                    isGUIDMaintain = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1993");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1993
            // If every EphemeralEntryID's ProviderUID equals server's Guid in NspiBind, it means the server maintain this GUID.
            Site.CaptureRequirementIfIsTrue(
                isGUIDMaintain,
                1993,
                @"[In Initialization] The server maintains the GUID used to identify an NSPI session during one call.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1725");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1725
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagEntryId,
                rowsOfQueryRows.Value.ARow[0].LpProps[0].PropTag,
                1725,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagEntryId, which is defined in [MS-OXOABK] section 2.2.3.2.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6100,
                @"[In PidTagEntryId] Property ID: 0x0FFF.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6101,
                @"[In PidTagEntryId] Data type: PtypBinary, 0x0102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1734");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1734
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagDisplayName,
                rowsOfQueryRows.Value.ARow[0].LpProps[1].PropTag,
                1734,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagDisplayName, which is defined in [MS-OXOABK] section 2.2.3.1.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6014,
                @"[In PidTagDisplayName] Property ID: 0x3001.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6015,
                @"[In PidTagDisplayName] Data type: Ptypstring, 0x001F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1731");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1731
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagDisplayType,
                rowsOfQueryRows.Value.ARow[0].LpProps[2].PropTag,
                1731,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagDisplayType, which is defined in [MS-OXOABK] section 2.2.3.11.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6035,
                @"[In PidTagDisplayType] Property ID: 0x3900.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6036,
                @"[In PidTagDisplayType] Data type: PtypInteger32, 0x0003.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1721");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1721
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagObjectType,
                rowsOfQueryRows.Value.ARow[0].LpProps[3].PropTag,
                1721,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagObjectType, which is defined in [MS-OXOABK] section 2.2.3.10.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                7109,
                @"[In PidTagObjectType] Property ID: 0x0FFE.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                7110,
                @"[In PidTagObjectType] Data type: PtypInteger32, 0x0003.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1722");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1722
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagInitialDetailsPane,
                rowsOfQueryRows.Value.ARow[0].LpProps[4].PropTag,
                1722,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagInitialDetailsPane, which is defined in [MS-OXOABK] section 2.2.3.33.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6605,
                @"[In PidTagInitialDetailsPane] Property ID: 0x3F08.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6606,
                @"[In PidTagInitialDetailsPane] Data type: PtypInteger32, 0x0003.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1723");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1723
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                rowsOfQueryRows.Value.ARow[0].LpProps[5].PropTag,
                1723,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagAddressBookDisplayNamePrintable, which is defined in [MS-OXOABK] section 2.2.3.7.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                4981,
                @"[In PidTagAddressBookDisplayNamePrintable] Property ID: 0x39FF.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                4982,
                @"[In PidTagAddressBookDisplayNamePrintable] Data type: Ptypstring, 0x001F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1724");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1724
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressBookContainerId,
                rowsOfQueryRows.Value.ARow[0].LpProps[6].PropTag,
                1724,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagAddressBookContainerId, which is defined in [MS-OXOABK] section 2.2.2.3.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                4969,
                @"[In PidTagAddressBookContainerId] Property ID: 0xFFFD.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                4970,
                @"[In PidTagAddressBookContainerId] Data type: PtypInteger32, 0x0003.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1726");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1726
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagInstanceKey,
                rowsOfQueryRows.Value.ARow[0].LpProps[7].PropTag,
                1726,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagInstanceKey, which is defined in [MS-OXOABK] section 2.2.3.6.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6626,
                @"[In PidTagInstanceKey] Property ID: 0x0FF6.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6627,
                @"[In PidTagInstanceKey] Data type: PtypBinary, 0x0102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1727");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1727
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagSearchKey,
                rowsOfQueryRows.Value.ARow[0].LpProps[8].PropTag,
                1727,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagSearchKey, which is defined in [MS-OXOABK] section 2.2.3.5.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                8385,
                @"[In PidTagSearchKey] Property ID: 0x300B");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                8386,
                @"[In PidTagSearchKey] Data type: PtypBinary, 0x0102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1728");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1728
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagRecordKey,
                rowsOfQueryRows.Value.ARow[0].LpProps[9].PropTag,
                1728,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagRecordKey, which is defined in [MS-OXOABK] section 2.2.3.4.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                7825,
                @"[In PidTagRecordKey] Property ID: 0x0FF9.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                7826,
                @"[In PidTagRecordKey] Data type: PtypBinary, 0x0102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1729");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1729
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressType,
                rowsOfQueryRows.Value.ARow[0].LpProps[10].PropTag,
                1729,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagAddressType, which is defined in [MS-OXOABK] section 2.2.3.13.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5316,
                @"[In PidTagAddressType] Property ID: 0x3002");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5317,
                @"[In PidTagAddressType] Data type: Ptypstring, 0x001F");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1730");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1730
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagEmailAddress,
                rowsOfQueryRows.Value.ARow[0].LpProps[11].PropTag,
                1730,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagEmailAddress, which is defined in [MS-OXOABK] section 2.2.3.14.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6056,
                @"[In PidTagEmailAddress] Property ID: 0x3003");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6057,
                @"[In PidTagEmailAddress] Data type: Ptypstring, 0x001F");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1732");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1732
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagTemplateid,
                rowsOfQueryRows.Value.ARow[0].LpProps[12].PropTag,
                1732,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagTemplateid, which is defined in [MS-OXOABK] section 2.2.3.3.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                8770,
                @"[In PidTagTemplateid] Property ID: 0x3902.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                8771,
                @"[In PidTagTemplateid] Data type: PtypBinary, 0x0102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1733");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1733
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagTransmittableDisplayName,
                rowsOfQueryRows.Value.ARow[0].LpProps[13].PropTag,
                1733,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagTransmittableDisplayName, which is defined in [MS-OXOABK] section 2.2.3.8.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                8815,
                @"[In PidTagTransmittableDisplayName] Property ID: 0x3A20.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                8816,
                @"[In PidTagTransmittableDisplayName] Data type: Ptypstring, 0x001F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1735");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1735
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagMappingSignature,
                rowsOfQueryRows.Value.ARow[0].LpProps[14].PropTag,
                1735,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagMappingSignature, which is defined in [MS-OXOABK] section 2.2.3.32.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6889,
                @"[In PidTagMappingSignature] Description: A 16-byte constant that is present on all Address Book objects.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6890,
                @"[In PidTagMappingSignature] Property ID: 0x0FF8.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                6891,
                @"[In PidTagMappingSignature] Data type: PtypBinary, 0x0102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1736");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1736
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressBookObjectDistinguishedName,
                rowsOfQueryRows.Value.ARow[0].LpProps[15].PropTag,
                1736,
                @"[In Required Properties] For every object in the address book, the server MUST minimally maintain the property: PidTagAddressBookObjectDistinguishedName, which is defined in [MS-OXOABK] section 2.2.3.15.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5163,
                @"[In PidTagAddressBookObjectDistinguishedName] Property ID: 0x803C.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5164,
                @"[In PidTagAddressBookObjectDistinguishedName] Data type: Ptypstring, 0x001F.");

            #endregion

            #endregion

            #region Call NspiQueryRows with SortType field of STAT block set to SortTypeDisplayName.
            count = 3; // Set the expected row count to be returned to a value greater than 1 to check the sort result.
            PropertyTagArray_r propTagsInstance2 = new PropertyTagArray_r
            {
                CValues = 2,
                AulPropTag = new uint[2]
                {
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagDisplayName,
                }
            };
            propTags = propTagsInstance2;
            stat.InitiateStat();
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R93.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R93
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsSortByName(rowsOfQueryRows, 1),
                93,
                @"[In Table Sort Orders] SortTypeDisplayName (0x00000000): The table is sorted ascending on the PidTagDisplayName property, as specified in [MS-OXCFOLD] section 2.2.2.2.2.5.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R967");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R967
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsSortByName(rowsOfQueryRows, 1),
                967,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 11] The server MUST return rows in the order they exist in the table being used.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1763");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1763
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsSortByName(rowsOfQueryRows, 1),
                1763,
                @"[In Absolute Positioning] The server sorts the objects in the address book container specified by ContainerID by the sort type specified in the SortType field and the default LCID NSPI_DEFAULT_LOCALE.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R997");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R997
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                997,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 17] If no other return values have been specified by these constraints [constraints 1-16], the server MUST return the return value ""Success"".");

            this.VerifyRowsReturnedFromNspiQueryRowsIsNotNull(rowsOfQueryRows);

            this.VerifyServerUsedTheListInNspiQueryRows(rowsOfQueryRows, propTags);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R962: The row count returned by the server is {0}.", rowsOfQueryRows.Value.CRows);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R962
            Site.CaptureRequirementIfIsTrue(
                rowsOfQueryRows.Value.CRows >= 1,
                962,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 10] If there are any rows that satisfy the client's query, the server MUST return at least one row.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiQueryRows operation with null proptags and non-null proptags.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC13_QueryRowsSuccessWithDifferentPropTags()
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

            #region Call NspiQueryRows with proptags not set to null.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fSkipObjects;
            uint tableCount = 0;
            uint[] table = null;
            uint count = Constants.QueryRowsRequestedRowNumber;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagPrimaryTelephoneNumber
                }
            };
            PropertyTagArray_r? propTags = propTagsInstance;

            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success.");
            #endregion

            #region Call NspiQueryRows method with proptags set to null.
            propTags = null;
            stat.InitiateStat();
            count = Constants.QueryRowsRequestedRowNumber;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success.");
            Site.Assert.IsNotNull(rowsOfQueryRows.Value.ARow, "The returned rows should not be empty. The row number is {0}.", rowsOfQueryRows == null ? 0 : rowsOfQueryRows.Value.CRows);

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1627");

            PropertyTagArray_r tempPropTags = new PropertyTagArray_r
            {
                CValues = 7,
                AulPropTag = new uint[7]
                {
                    (uint)AulProp.PidTagAddressBookContainerId,
                    (uint)AulProp.PidTagObjectType,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagPrimaryTelephoneNumber,
                    (uint)AulProp.PidTagDepartmentName,
                    (uint)AulProp.PidTagOfficeLocation,
                }
            };
            PropertyTagArray_r? defaultPropTags = tempPropTags;

            bool isPropTagConsistency = true;

            foreach (PropertyRow_r propertyrow in rowsOfQueryRows.Value.ARow)
            {
                for (int i = 0; i < defaultPropTags.Value.CValues; i++)
                {
                    if ((defaultPropTags.Value.AulPropTag[i] & 0xffff0000) != (propertyrow.LpProps[i].PropTag & 0xffff0000))
                    {
                        Site.Log.Add(
                            LogEntryKind.Debug,
                            "The default property tag of index {0} is {1}, the server returned property tag in property row is {2}",
                            i,
                            defaultPropTags.Value.AulPropTag[i],
                            propertyrow.LpProps[i].PropTag);

                        isPropTagConsistency = false;
                        break;
                    }
                }

                if (isPropTagConsistency == false)
                {
                    break;
                }
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1627
            Site.CaptureRequirementIfIsTrue(
                isPropTagConsistency,
                1627,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] [If the input parameter pPropTags is NULL] This server MUST use the following proptag list (using proptags defined in [MS-OXPROPS]), in this order: {PidTagAddressBookContainerId ([MS-OXOABK] section 2.2.2.3), PidTagObjectType ([MS-OXOABK] section 2.2.3.10), PidTagDisplayType ([MS-OXOABK] section 2.2.3.11), PidTagDisplayName ([MS-OXOABK] section 2.2.3.1) with the property type PtypString8, as specified in [MS-OXCDATA] section 2.11.1, PidTagPrimaryTelephoneNumber ([MS-OXOCNTC] section 2.2.1.4.5) with the property type PtypString8, PidTagDepartmentName ([MS-OXOABK] section 2.2.4.6) with the property type PtypString8, PidTagOfficeLocation ([MS-OXOABK] section 2.2.4.5) with the property type PtypString8}");

            this.VerifyRowsReturnedFromNspiQueryRowsIsNotNull(rowsOfQueryRows);

            string userName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
            PropertyRow_r? userProperties = null;
            foreach (PropertyRow_r propRow in rowsOfQueryRows.Value.ARow)
            {
                string displayName = System.Text.Encoding.UTF8.GetString(propRow.LpProps[3].Value.LpszA);
                if (displayName.Equals(userName, StringComparison.OrdinalIgnoreCase))
                {
                    userProperties = propRow;
                    break;
                }
            }

            Site.Assert.IsNotNull(userProperties, "The properties of the address object with display name {0} should be found.", userName);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7545");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R7545
            Site.CaptureRequirementIfAreEqual<uint>(
                0x3A1A0000,
                userProperties.Value.LpProps[4].PropTag & 0xffff0000,
                "MS-OXPROPS",
                7545,
                @"[In PidTagPrimaryTelephoneNumber] Property ID: 0x3A1A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5986");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R5986
            Site.CaptureRequirementIfAreEqual<uint>(
                0x3a180000,
                userProperties.Value.LpProps[5].PropTag & 0xffff0000,
                "MS-OXPROPS",
                5986,
                @"[In PidTagDepartmentName] Property ID: 0x3A18.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7116");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R7116
            Site.CaptureRequirementIfAreEqual<uint>(
                0x3A190000,
                userProperties.Value.LpProps[6].PropTag & 0xffff0000,
                "MS-OXPROPS",
                7116,
                @"[In PidTagOfficeLocation] Property ID: 0x3A19.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiQueryRows operation with specified proptags.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC14_QueryRowsSuccessWithSpeicfiedPropTag()
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

            #region Call NspiQueryRows with proptags containing PidTagContainerContents and PidTagContainerFlags.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            uint tableCount = 0;
            uint[] table = null;
            uint count = Constants.QueryRowsRequestedRowNumber;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 5,
                AulPropTag = new uint[5]
                {
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType,
                    (uint)AulProp.PidTagContainerContents,
                    (uint)AulProp.PidTagContainerFlags,
                }
            };
            if (this.Transport == "mapi_http")
            {
                propTagsInstance = new PropertyTagArray_r
                {
                    CValues = 4,
                    AulPropTag = new uint[4]
                    {
                        (uint)AulProp.PidTagDisplayName,
                        (uint)AulProp.PidTagDisplayType,
                        (uint)AulProp.PidTagObjectType,
                        (uint)AulProp.PidTagContainerFlags
                    }
                };
            }

            PropertyTagArray_r? propTags = propTagsInstance;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success.");

            #region Capture code
            string distributeListName = Common.GetConfigurationPropertyValue("DistributionListName", this.Site);

            // A variable used to record the PidTagContainerContents propTag got from server.
            uint actualPidTagContainerContents = 0;

            // A variable used to record the PidTagContainerFlags propTag got from server.
            uint actualPidTagContainerFlags = 0;

            // A variable used to record the PidTagObjectType propTag got from server.
            int actualPidTagObjectType = 0;

            for (int i = 0; i < rowsOfQueryRows.Value.CRows; i++)
            {
                string name = System.Text.Encoding.UTF8.GetString(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(distributeListName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
                    {
                        actualPidTagObjectType = rowsOfQueryRows.Value.ARow[i].LpProps[2].Value.L;
                        actualPidTagContainerContents = rowsOfQueryRows.Value.ARow[i].LpProps[3].PropTag;
                        actualPidTagContainerFlags = rowsOfQueryRows.Value.ARow[i].LpProps[4].PropTag;

                        break;
                    }
                    else
                    {
                        actualPidTagObjectType = rowsOfQueryRows.Value.ARow[i].LpProps[2].Value.L;
                        actualPidTagContainerFlags = rowsOfQueryRows.Value.ARow[i].LpProps[3].PropTag;
                    }
                }
            }

            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1737");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1737
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)AulProp.PidTagContainerContents,
                    actualPidTagContainerContents,
                    1737,
                    @"[In Required Properties] The server MUST maintain the property PidTagContainerContents, which is defined in [MS-OXOABK] [MS-OXOABK] section 2.2.6.3, and has a PidTagObjectType property with the DISTLIST value, as specified in [MS-OXOABK] section 2.2.3.10.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1738");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1738
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)AulProp.PidTagContainerFlags,
                    actualPidTagContainerFlags,
                    1738,
                    @"[In Required Properties] The server MUST maintain the property PidTagContainerFlags, which is defined in [MS-OXOABK] section 2.2.2.1, and has a PidTagObjectType property with the DISTLIST value, as specified in [MS-OXOABK] section 2.2.3.10.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXOABK_R892");

            // Verify MS-OXOABK requirement: MS-OXOABK_R892
            Site.CaptureRequirementIfAreEqual<int>(
                0x00000008,
                actualPidTagObjectType,
                "MS-OXOABK",
                892,
                @"[In PidTagObjectType] If the Address Book object is DISTLIST, the value of the PidTagObjectType property is 0x00000008.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5798,
                @"[In PidTagContainerContents] Description: Always empty. An NSPI server defines this value for distribution lists and it is not 
                        present for other objects.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5799,
                @"[In PidTagContainerContents] Property ID: 0x360F.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5800,
                @"[In PidTagContainerContents] Data type: PtypEmbeddedTable, 0x000D");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5806,
                @"[In PidTagContainerFlags] Property ID: 0x3600.");

            // If the codes can reach here, the requirement based on this must have been captured, so it can be captured directly.
            Site.CaptureRequirement(
                "MS-OXPROPS",
                5807,
                @"[In PidTagContainerFlags] Data type: PtypInteger32, 0x0003.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiQueryRows operation with non-null table.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC15_QueryRowsSuccessWithNotNullTable()
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

            #region Call NspiGetMatches method to get a list of valid Minimal Entry IDs.
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

            PropertyTagArray_r propTags1 = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags1;

            // Set value for propNameOfGetMatches.
            PropertyName_r? propNameOfGetMatches = null;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            Site.Assert.IsTrue(rowsOfGetMatches != null && outMIds != null, "NspiGetMatches should return valid result used for NspiQueryRows!");
            #endregion

            #region Call NspiQueryRows method using the output parameter outMids of NspiGetMatches as the input parameter lpETable.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            uint tableCount = outMIds.Value.CValues;
            uint[] table = new uint[outMIds.Value.CValues];
            Array.Copy(outMIds.Value.AulPropTag, table, outMIds.Value.CValues);
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
                    (uint)AulProp.PidTagObjectType
                }
            };
            PropertyTagArray_r? propTags = propTagsInstance;

            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R974");

            bool isCorresponds = true;
            for (int i = 0; i < table.Length; i++)
            {
                if (table[i] != AdapterHelper.ParseEphemeralEntryIDFromBytes(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.Bin.Lpb).Mid)
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "The row of index {0} in RowSet does not correspond to a row in the table. The row in RowSet is {1}, the row in table is {2}.",
                        i,
                        AdapterHelper.ParseEphemeralEntryIDFromBytes(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.Bin.Lpb).Mid,
                        table[i]);

                    isCorresponds = false;
                    break;
                }
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R974
            Site.CaptureRequirementIfIsTrue(
                isCorresponds,
                974,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] Each row in the RowSet corresponds to a row in the table specified by input parameters.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R975");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R975
            Site.CaptureRequirementIfIsTrue(
                isCorresponds,
                975,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] The rows in the RowSet are in a one-to-one order preserving correspondence with the rows in the table specified by input parameters.");

            string mailUserName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
            string distributionListName = Common.GetConfigurationPropertyValue("DistributionListName", this.Site);
            string forumName = Common.GetConfigurationPropertyValue("ForumName", this.Site);
            string agentName = Common.GetConfigurationPropertyValue("AgentName", this.Site);
            string remoteMailUserName = Common.GetConfigurationPropertyValue("RemoteMailUserName", this.Site);

            bool isDT_MAILUSERTypeCorrect = false;
            bool isDT_DISTLISTTypeCorrect = false;
            bool isDT_FORUMTypeCorrect = false;
            bool isDT_AGENTTypeCorrect = false;
            bool isDT_REMOTE_MAILUSERTypeCorrect = false;

            string name = string.Empty;
            DisplayTypeValue displayType = DisplayTypeValue.DT_SEARCH;

            foreach (PropertyRow_r row in rowsOfQueryRows.Value.ARow)
            {
                // Get name and the corresponding display type.
                foreach (PropertyValue_r val in row.LpProps)
                {
                    if (val.PropTag == (uint)AulProp.PidTagDisplayName)
                    {
                        name = System.Text.Encoding.UTF8.GetString(val.Value.LpszA);
                    }
                    else if (val.PropTag == (uint)AulProp.PidTagDisplayType)
                    {
                        displayType = (DisplayTypeValue)val.Value.L;
                    }
                }

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(mailUserName.ToLower(System.Globalization.CultureInfo.CurrentCulture)) && displayType == DisplayTypeValue.DT_MAILUSER)
                {
                    isDT_MAILUSERTypeCorrect = true;
                }

                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(distributionListName.ToLower(System.Globalization.CultureInfo.CurrentCulture)) && displayType == DisplayTypeValue.DT_DISTLIST)
                {
                    isDT_DISTLISTTypeCorrect = true;
                }

                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(forumName.ToLower(System.Globalization.CultureInfo.CurrentCulture)) && displayType == DisplayTypeValue.DT_FORUM)
                {
                    isDT_FORUMTypeCorrect = true;
                }

                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(agentName.ToLower(System.Globalization.CultureInfo.CurrentCulture)) && displayType == DisplayTypeValue.DT_AGENT)
                {
                    isDT_AGENTTypeCorrect = true;
                }

                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(remoteMailUserName.ToLower(System.Globalization.CultureInfo.CurrentCulture)) && displayType == DisplayTypeValue.DT_REMOTE_MAILUSER)
                {
                    isDT_REMOTE_MAILUSERTypeCorrect = true;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R48: the display name is {0}, display type is {1}", name, displayType);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R48
            Site.CaptureRequirementIfIsTrue(
                isDT_MAILUSERTypeCorrect,
                48,
                @"[In Display Type Values] DT_MAILUSER display type name with 0x00000000 means A typical messaging user.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R49: the display name is {0}, display type is {1}", name, displayType);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R49
            Site.CaptureRequirementIfIsTrue(
                isDT_DISTLISTTypeCorrect,
                49,
                @"[In Display Type Values] DT_DISTLIST display type name with 0x00000001 means A distribution list.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R50: the display name is {0}, display type is {1}", name, displayType);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R50
            Site.CaptureRequirementIfIsTrue(
                isDT_FORUMTypeCorrect,
                50,
                @"[In Display Type Values] DT_FORUM display type with 0x00000002 value means A forum, such as a bulletin board service or a public or shared folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R51: the display name is {0}, display type is {1}", name, displayType);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R51
            Site.CaptureRequirementIfIsTrue(
                isDT_AGENTTypeCorrect,
                51,
                @"[In Display Type Values] DT_AGENT display type with 0x00000003 value means An automated agent, such as Quote-Of-The-Day or a weather chart display.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R55: the display name is {0}, display type is {1}", name, displayType);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R55
            Site.CaptureRequirementIfIsTrue(
                isDT_REMOTE_MAILUSERTypeCorrect,
                55,
                @"[In Display Type Values] DT_REMOTE_MAILUSER display type with 0x00000006 value means An Address Book object known to be from a foreign or remote messaging system.");

            this.VerifyRowsReturnedFromNspiQueryRowsIsNotNull(rowsOfQueryRows);

            this.VerifyServerUsedTheListInNspiQueryRows(rowsOfQueryRows, propTags);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R958: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
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

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R958
            bool isVerifyR958 = inputStat.CodePage == stat.CodePage
                                || inputStat.ContainerID == stat.ContainerID
                                || inputStat.CurrentRec == stat.CurrentRec
                                || inputStat.Delta == stat.Delta
                                || inputStat.NumPos == stat.NumPos
                                || inputStat.SortLocale == stat.SortLocale
                                || inputStat.SortType == stat.SortType
                                || inputStat.TemplateLocale == stat.TemplateLocale
                                || inputStat.TotalRecs == stat.TotalRecs;

            // According to the description in section 3.1.4.1.8, if the server is using the table specified by the input parameter pStat, the server must update the status of the table. 
            // So if the returned STAT is not updated, it is believed that the server uses the table specified by the input parameter lpETable.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR958,
                958,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] [If the input parameter lpETable is not NULL] The server MUST use that table when constructing the return parameter ppRows.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1077
            // If the test case can reach here, it is believed that the mids returned by NspiGetMatches are used for NspiQueryRows to specify an Explicit Table, so this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1077,
                @"[In NspiGetMatches] The NspiGetMatches method returns an Explicit Table.");
            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to ignoring other values of dwFlags that are not fEphID or fSkipObjects
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC16_QueryRowsIgnoreSomeFlags()
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

            #region Call NspiQueryRows with dwFlags set to one value that is not fEphID (0x2) or fSkipObjects (0x1).
            uint flagsOfQueryRows = 3;
            uint tableCount = 0;
            uint[] table = null;
            uint count = 10;
            stat.InitiateStat();
            PropertyRowSet_r? rowsOfQueryRows1;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 4,
                AulPropTag = new uint[4]
                {
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType,
                    (uint)AulProp.PidTagContainerFlags,
                }
            };
            PropertyTagArray_r? propTags = propTagsInstance;
            ErrorCodeValue result1 = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result1, "NspiQueryRows should return Success");
            #endregion

            #region Call NspiQueryRows with dwFlags set to another value that is not fEphID (0x2) or fSkipObjects (0x1).
            flagsOfQueryRows = 4;
            stat.InitiateStat();
            PropertyRowSet_r? rowsOfQueryRows2;
            ErrorCodeValue result2 = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows2);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result2, "NspiQueryRows should return Success");

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1926", "The rows number of the first call is {0}, the rows number of the second call is {1}.", rowsOfQueryRows1.Value.CRows, rowsOfQueryRows2.Value.CRows);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1926
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowSetEqual(rowsOfQueryRows1, rowsOfQueryRows2),
                1926,
                @"[In NspiQueryRows] If dwflags is set to different values other than the bit flags fEphID and fSkipObjects, server will return the same result.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to comparing NspiQueryRows and NspiGetProps operations.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC17_QueryRowsUseSameParametersWithNspiGetProps()
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

            #region Call NspiUpdateStat to update the STAT block that represents position in a table to reflect positioning changes requested by the client.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiQueryRows with the specified proptags.
            STAT savedStat = stat;
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

            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success.");

            #region Capture
            this.VerifyRowsReturnedFromNspiQueryRowsIsNotNull(rowsOfQueryRows);

            this.VerifyServerUsedTheListInNspiQueryRows(rowsOfQueryRows, propTags);
            #endregion Capture
            #endregion

            #region Call NspiGetProps using the same parameters with those of NspiQueryRows.
            PropertyRow_r? rows;
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfQueryRows, savedStat, propTags, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetProps should return Success.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1694");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1694
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowEqual(rowsOfQueryRows.Value.ARow[rowsOfQueryRows.Value.CRows - 1], rows),
                1694,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] The Rows placed into the RowSet are exactly those rows that would be returned to the client in the ppRows output parameter of the NspiGetProps method, as specified in section 3.1.4.1.7, using the following parameters:
                  The NspiQueryRows parameter hRpc is used as the NspiGetProps parameter hRpc. 
                  The NspiQueryRows parameter dwFlags is used as the NspiGetProps parameter dwFlags. 
                  The NspiQueryRows input parameter pStat is used as the NspiGetProps parameter pStat. The CurrentRec field is set to the Minimal Entry ID of the row being returned.
                  The list of proptags the server constructs as specified by constraint 6 is used as the NspiGetProps parameter pPropTags.");

            this.VerifyRowsReturnedFromNspiQueryRowsIsNotNull(rowsOfQueryRows);

            this.VerifyServerUsedTheListInNspiQueryRows(rowsOfQueryRows, propTags);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R983");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R983
            // This value is returned from the server, if it is not null, it illustrates that the server returned the structure.
            Site.CaptureRequirementIfIsNotNull(
                rows,
                983,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] Otherwise [If a call to the NspiGetProps method with parameters hRpc, dwFlags, pStat and pPropTags would return ""Success"" or ""ErrorsReturned""], the server MUST return the RowSet constructed in the output parameter ppRows.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to comparing NspiQueryRows and NspiUpdateStat operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC18_QueryRowsUpdateStatCompareNspiUpdateStat()
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

            #region Call NspiQueryRows method to query rows which contain the specific properties.
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

            STAT backStat = stat;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success");

            #region Capture

            this.VerifyRowsReturnedFromNspiQueryRowsIsNotNull(rowsOfQueryRows);

            this.VerifyServerUsedTheListInNspiQueryRows(rowsOfQueryRows, propTags);

            #endregion Capture
            #endregion

            #region Call NspiUpdateStat method with the delta field set to null.
            uint reserved = 0;
            int? delta = null;
            STAT updateStat = backStat;
            updateStat.Delta += (int)rowsOfQueryRows.Value.CRows;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref updateStat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success.");

            #region Capture
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_1701: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
                "the SortType of updateStat is {9}, the ContainerID of updateStat is {10}, the CurrentRec of updateStat is {11}, the Delta of updateStat is {12}, the NumPos of updateStat is {13}, the TotalRecs of updateStat is {13}, the CodePage of updateStat is {14}, the TemplateLocale of updateStat is {15}, the SortLocale of updateStat is {16}",
                stat.SortType,
                stat.ContainerID,
                stat.CurrentRec,
                stat.Delta,
                stat.NumPos,
                stat.TotalRecs,
                stat.CodePage,
                stat.TemplateLocale,
                stat.SortLocale,
                updateStat.SortType,
                updateStat.ContainerID,
                updateStat.CurrentRec,
                updateStat.Delta,
                updateStat.NumPos,
                updateStat.TotalRecs,
                updateStat.CodePage,
                updateStat.TemplateLocale,
                updateStat.SortLocale);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1701
            bool isVerifyR1701 = (stat.SortType == updateStat.SortType)
                                && (stat.ContainerID == updateStat.ContainerID)
                                && (stat.CurrentRec == updateStat.CurrentRec)
                                && (stat.Delta == updateStat.Delta)
                                && (stat.NumPos == updateStat.NumPos)
                                && (stat.TotalRecs == updateStat.TotalRecs)
                                && (stat.CodePage == updateStat.CodePage)
                                && (stat.TemplateLocale == updateStat.TemplateLocale)
                                && (stat.SortLocale == updateStat.SortLocale);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1701,
                1701,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 16] This update [update the status of the table] MUST be exactly the same update that would occur via the NspiUpdateStat method with the following parameters:
                 The NspiQueryRows parameter hRpc is used as the NspiUpdateStat parameter hRpc.
                 The value 0 is used as NspiUpdateStat parameter Reserved. 
                 The NspiQueryRows output parameter pStat (as modified by the prior constraints) is used as the NSPIUpdateStat parameter pStat. The number of rows returned in the NspiQueryRows output parameter ppRows is added to the Delta field.
                 The value NULL is used as the NspiUpdateStat parameter plDelta. ");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to different table sort orders.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC19_GetMatchesWithDifferentSortOrder()
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

            #region Call NspiGetMatches method with SortType field of STAT block set to SortTypeDisplayName_RO.
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName_RO;
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;
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
                AulPropTag = new uint[]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                    (uint)AulProp.PidTagDisplayType
                }
            };
            propTags.CValues = (uint)propTags.AulPropTag.Length;
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success");

            #region Capture code

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R99");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R99
            // The field LpszW indicates a single Unicode string value.
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                99,
                "[In Table Sort Orders] [SortTypeDisplayName_RO] The client MUST set this value only when using the NspiGetMatches method, as specified in section 3.1.4.1.10, to open a nonwritable table on an object-valued property.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R98");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R98
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsSortByName(rowsOfGetMatches, 1),
                98,
                @"[In Table Sort Orders] SortTypeDisplayName_RO (0x000003E8): The table is sorted ascending on the PidTagDisplayName property.");
            #endregion
            #endregion

            #region Call NspiGetMatches method with SortType field of STAT block set to SortTypeDisplayName_W.
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName_W;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);

            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R100");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R100
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsSortByName(rowsOfGetMatches, 1),
                100,
                @"[In Table Sort Orders] SortTypeDisplayName_W (0x000003E9): The table is sorted ascending on the PidTagDisplayName property.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R101");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R101
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                101,
                "[In Table Sort Orders] [SortTypeDisplayName_W] The client MUST set this value only when using the NspiGetMatches method to open a writable table on an object-valued property.");
            #endregion
            #endregion

            #region Call NspiGetMatches method with specified filter and prop tags.
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
                                PropTag = (uint)AulProp.PidTagAddressBookPhoneticDisplayName
                            }
                    }
            };
            filter = res_r;

            PropertyTagArray_r prop = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagAddressBookPhoneticDisplayName
                }
            };

            propTagsOfGetMatches = prop;
            propNameOfGetMatches = null;

            stat.SortType = (uint)TableSortOrder.SortTypePhoneticDisplayName;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success.");

            #region Capture

            // Whether the row is sorted by display name.
            bool isSortyByPhoneticDisplayName = false;

            // Store name of each row.
            string[] phoneticDisplayName = new string[rowsOfGetMatches.Value.CRows];
            Site.Log.Add(LogEntryKind.Debug, "The rows sorted by phonetic display name are:");
            for (int i = 0; i < rowsOfGetMatches.Value.CRows; i++)
            {
                phoneticDisplayName[i] = System.Text.Encoding.UTF8.GetString(rowsOfGetMatches.Value.ARow[i].LpProps[1].Value.LpszA);
                Site.Log.Add(LogEntryKind.Debug, "Row {0}: {1}", i, phoneticDisplayName[i]);
            }

            for (int i = 0; i < phoneticDisplayName.Length - 1; i++)
            {
                if (string.Compare(phoneticDisplayName[i], phoneticDisplayName[i + 1], StringComparison.Ordinal) < 0)
                {
                    isSortyByPhoneticDisplayName = true;
                }
                else
                {
                    isSortyByPhoneticDisplayName = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R95");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R95
            Site.CaptureRequirementIfIsTrue(
                isSortyByPhoneticDisplayName,
                95,
                @"[In Table Sort Orders] SortTypePhoneticDisplayName (0x00000003): The table is sorted ascending on the PidTagAddressBookPhoneticDisplayName property, as specified in [MS-OXOABK] section 2.2.3.9.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R495");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R495
            this.Site.CaptureRequirementIfIsTrue(
                isSortyByPhoneticDisplayName,
                495,
                @"[In String Sorting] If the server supports the SortTypePhoneticDisplayName sort order, it [server] MUST also support sorting on Unicode string representation for the PidTagAddressBookPhoneticDisplayName property.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetMatches operation returning success.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC20_GetMatchesSuccess()
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

            #region Call NspiQueryRows with fEphID dwFlags to get an EphemeralEntryID.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            uint tableCount = 0;
            uint[] table = null;
            uint count = 10;
            PropertyRowSet_r? rowsOfQueryRows1;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 2,
                AulPropTag = new uint[2]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                }
            };
            PropertyTagArray_r? propTags1 = propTagsInstance;

            stat.CodePage = (uint)RequiredCodePage.CP_TELETEX;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags1, out rowsOfQueryRows1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");
            Site.Assert.IsNotNull(rowsOfQueryRows1.Value.ARow, "The returned rows should not be empty. The row number is {0}.", rowsOfQueryRows1 == null ? 0 : rowsOfQueryRows1.Value.CRows);
            #endregion

            #region Call NspiUpdateStat to update STAT block.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiGetMatches method with the specified filter and proptags.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            Restriction_r entryIdExist = new Restriction_r
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
            Restriction_r? filter = entryIdExist;

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

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1155");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1155
            // The parameter rowsOfGetMatches which contains a PropertyRowSet_r structure is returned from the server, if this value is not null, 
            // it illustrates  that the server must have constructed the structure.
            Site.CaptureRequirementIfIsNotNull(
                rowsOfGetMatches,
                1155,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 18] Subject to the prior constraints, the server MUST construct a PropertyRowSet_r to return to the client in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1167");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1167
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1167,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 10] If no other return values have been specified by these constraints [constraints 1-9], the server MUST return the return value ""Success"".");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1090, the maximum number of rows to return in a restricted address book container is {0}, the actual number is {1}",
                requested,
                rowsOfGetMatches.Value.CRows);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1090
            this.Site.CaptureRequirementIfIsTrue(
                rowsOfGetMatches.Value.CRows <= requested,
                1090,
                @"[In NspiGetMatches] [ulRequested] Contains the maximum number of rows to return in a restricted address book container.");

            int index = 0;
            foreach (PropertyRow_r arow in rowsOfGetMatches.Value.ARow)
            {
                EphemeralEntryID temp = AdapterHelper.ParseEphemeralEntryIDFromBytes(arow.LpProps[0].Value.Bin.Lpb);
                Site.Assert.AreEqual<uint>(outMIds.Value.AulPropTag[index], temp.Mid, "The Minimal Entry ID hold by ppOutMIds should be same with the MID in the EphemeralEntryID.");
                index++;
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1092
            // All the Minimal Entry IDs contained in ppOutMIds had been checked, MS-OXNSPI_R1092 can be verified directly.
            this.Site.CaptureRequirement(
                1092,
                @"[In NspiGetMatches] [ppOutMIds] On return, it holds a list of Minimal Entry IDs that comprise a restricted address book container.");

            PropertyRow_r rowValue = rowsOfGetMatches.Value.ARow[0];

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1094, the first property tag value of row value is {0}, the second is {1}, the third is {2}",
                rowValue.LpProps[0].PropTag,
                rowValue.LpProps[1].PropTag,
                rowValue.LpProps[2].PropTag);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1094
            bool isVerifiedR1094 = rowValue.LpProps[0].PropTag.Equals((uint)AulProp.PidTagEntryId) &&
                                   rowValue.LpProps[1].PropTag.Equals((uint)AulProp.PidTagDisplayType) &&
                                   rowValue.LpProps[2].PropTag.Equals((uint)AulProp.PidTagObjectType);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1094,
                1094,
                @"[In NspiGetMatches] [pPropTags] Contains list of the proptags of the columns that client wants to be returned for each row returned.");

            // pStat is defined to the type of STAT in this test suite, and the CurrentRec describes the logical position. 
            // So this requirement can be captured if code can reach here.
            this.Site.CaptureRequirement(
                1081,
                @"[In NspiGetMatches] pStat: A reference to a STAT block describing a logical position in a specific address book container.");

            #endregion
            #endregion

            #region Call NspiGetMatches method with another Reserved2 value.
            uint reserved2WithAnotherValue = 1;

            // Output parameters.
            PropertyTagArray_r? outMIdsWithAnotherReserved2Value;
            PropertyRowSet_r? rowsOfGetMatchesWithAnotherReserved2Value;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserved2WithAnotherValue, filter, propNameOfGetMatches, requested, out outMIdsWithAnotherReserved2Value, propTagsOfGetMatches, out rowsOfGetMatchesWithAnotherReserved2Value);

            #region Capture

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1899
            bool isMidEqual = true;
            if (outMIds.Value.CValues == outMIdsWithAnotherReserved2Value.Value.CValues)
            {
                for (int i = 0; i < outMIds.Value.CValues; i++)
                {
                    if (outMIds.Value.AulPropTag[i] != outMIdsWithAnotherReserved2Value.Value.AulPropTag[i])
                    {
                        isMidEqual = false;
                        break;
                    }
                }
            }
            else
            {
                isMidEqual = false;
            }

            bool isRowsOfGetMatchesEqual = true;
            if (rowsOfGetMatches.Value.CRows == rowsOfGetMatchesWithAnotherReserved2Value.Value.CRows)
            {
                for (int i = 0; i < rowsOfGetMatches.Value.CRows; i++)
                {
                    if (rowsOfGetMatches.Value.ARow[i].CValues == rowsOfGetMatchesWithAnotherReserved2Value.Value.ARow[i].CValues)
                    {
                        for (int j = 0; j < rowsOfGetMatches.Value.ARow[1].CValues; j++)
                        {
                            if (rowsOfGetMatches.Value.ARow[i].LpProps[j].PropTag != rowsOfGetMatchesWithAnotherReserved2Value.Value.ARow[i].LpProps[j].PropTag)
                            {
                                isRowsOfGetMatchesEqual = false;
                                break;
                            }
                        }
                    }
                    else
                    {
                        isRowsOfGetMatchesEqual = false;
                        break;
                    }
                }
            }
            else
            {
                isRowsOfGetMatchesEqual = false;
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1899.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1899
            this.Site.CaptureRequirementIfIsTrue(
                isMidEqual && isRowsOfGetMatchesEqual,
                1899,
                @"[In NspiGetMatches] If this field [Reserved2] is set to different values, the server will return the same result.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to comparing NspiGetMatches and NspiQueryRows operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC21_GetMatchesComparedWithNspiQueryRows()
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

            #region Call NspiGetMatches method to restrict a specific table based on the input parameters and return the resultant Explicit Table.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;
            PropertyName_r? propNameOfGetMatches = null;

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

            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                CValues = 4,
                AulPropTag = new uint[4]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayName,
                    (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                    (uint)AulProp.PidTagDisplayType
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success");
            #endregion

            #region Call NspiQueryRows method to compare with NspiGetMatches result.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            uint tableCount = outMIds.Value.CValues;
            uint[] table = new uint[outMIds.Value.CValues];
            Array.Copy(outMIds.Value.AulPropTag, table, outMIds.Value.CValues);

            uint count = outMIds.Value.CValues;
            PropertyRowSet_r? rowsOfQueryRows;
            STAT inputStat = stat;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1696");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1696
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowSetEqual(rowsOfGetMatches, rowsOfQueryRows),
                1696,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [constraint 18] This PropertyRowSet_r MUST be exactly the same PropertyRowSet_r that would be returned in the ppRows parameter of a call to the NspiQueryRows method with the following parameters:
                The NspiGetMatches parameter hRpc is used as the NspiQueryRows parameter hRpc. 
                The value ""fEphID"" is used as the NspiQueryRows parameter dwFlags.
                The NspiGetMatches output parameter pStat (as modified by the prior constraints) is used as the NspiQueryRows parameter pStat.
                The number of Minimal Entry IDs in the constructed Explicit Table is used as the NspiQueryRows parameter dwETableCount.
                The constructed Explicit Table is used as the NspiQueryRows parameter lpETable. These Minimal Entry IDs are expressed as a simple array of DWORD values rather than as a PropertyTagArray_r value.
                The number of Minimal Entry IDs in the constructed Explicit Table is used as the NspiQueryRows parameter Count.
                The NspiGetMatches parameter proptags is used as the NspiQueryRows parameter proptags.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1165: the SortType of output stat is {0}, the ContainerID of output stat is {1}, the CurrentRec of output stat is {2}, the Delta of output stat is {3}, the NumPos of output stat is {4}, the TotalRecs of output stat is {5}, the CodePage of output stat is {6}, the TemplateLocale of output stat is {7}, the SortLocale of output stat is {8};" +
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

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1165
            bool isVerifyR1165 = (stat.SortType == inputStat.SortType)
                                && (stat.ContainerID == inputStat.ContainerID)
                                && (stat.CurrentRec == inputStat.CurrentRec)
                                && (stat.Delta == inputStat.Delta)
                                && (stat.NumPos == inputStat.NumPos)
                                && (stat.TotalRecs == inputStat.TotalRecs)
                                && (stat.CodePage == inputStat.CodePage)
                                && (stat.TemplateLocale == inputStat.TemplateLocale)
                                && (stat.SortLocale == inputStat.SortLocale);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1165,
                1165,
                @"[In NspiGetMatches] Note that the server MUST NOT modify the return value of 
                the NspiGetMatches method output parameter pStat in any way in the process of constructing the output PropertyRowSet_r.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiResortRestriction operation returning success.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC22_ResortRestrictionSuccess()
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

            #region Call NspiDNToMId method to get valid MIDs.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                CValues = 3,
                LppszA = new string[3]
                {
                    Common.GetConfigurationPropertyValue("User2Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User1Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User3Essdn", this.Site),
                }
            };

            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiDNToMId should return Success!");
            #endregion

            #region Call NspiResortRestriction method with all valid parameters.
            uint reservedOfResortRestriction = 0;
            PropertyTagArray_r inmids = new PropertyTagArray_r();
            inmids = mids.Value;
            stat.InitiateStat();

            PropertyTagArray_r? outMIds = null;
            this.Result = this.ProtocolAdatper.NspiResortRestriction(reservedOfResortRestriction, ref stat, inmids, ref outMIds);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiResortRestriction method should succeed.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R538");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R538
            Site.CaptureRequirementIfAreEqual<uint>(
                stat.TotalRecs,
                outMIds.Value.CValues,
                538,
                @"[In Absolute Positioning] [step 3] The server reports this[the number of objects in the sorted table] in the TotalRecs field of the STAT structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1209");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1209
            Site.CaptureRequirementIfAreEqual(
                ErrorCodeValue.Success,
                this.Result,
                1209,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 9] If no other return values have been specified by these constraints [constraints 1-8], the server MUST return the return value ""Success"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R897");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R897
            Site.CaptureRequirementIfAreEqual(
                ErrorCodeValue.Success,
                this.Result,
                "MS-OXCDATA",
                897,
                @"[In Error Codes] Success(S_OK, SUCCESS_SUCCESS) will be returned, if the operation succeeded.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R898");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R898
            Site.CaptureRequirementIfAreEqual<uint>(
                0x0,
                (uint)this.Result,
                "MS-OXCDATA",
                898,
                @"[In Error Codes] The numeric value (hex) for error code Success is 0x00000000, %x00.00.00.00.");
            #endregion

            #endregion

            #region Call NspiQueryRows method to query rows which contain the specific properties.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fSkipObjects;
            uint tableCount = outMIds.Value.CValues;
            uint[] table = outMIds.Value.AulPropTag;
            uint count = outMIds.Value.CValues;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagDisplayName
                }
            };
            PropertyTagArray_r? propTags = propTagsInstance;

            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success");

            #region Capture

            // Add the debug information
            bool isCorrectOrder = false;
            for (int i = 0; i < count - 1; i++)
            {
                string displayName = System.Text.Encoding.Default.GetString(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.LpszA);
                string nextDisplayName = System.Text.Encoding.Default.GetString(rowsOfQueryRows.Value.ARow[i + 1].LpProps[0].Value.LpszA);
                if (i == 0)
                {
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1168. The display name of the No.{0} object in the table is {1}.", i, displayName);
                }

                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1168. The display name of the No.{0} object in the table is {1}.", i + 1, nextDisplayName);
                if (string.Compare(displayName, nextDisplayName, StringComparison.Ordinal) > 0)
                {
                    isCorrectOrder = false;
                    break;
                }
                else
                {
                    isCorrectOrder = true;
                }
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1168
            this.Site.CaptureRequirementIfIsTrue(
                isCorrectOrder,
                1168,
                @"[In NspiResortRestriction] The NspiResortRestriction method applies a sort order to the objects in a restricted address book container.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiResortRestriction constructs explicit table.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC23_ResortRestrictionSuccessCheckRowNumber()
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

            #region Call NspiGetMatches method to get a list of valid MIDs (Minimal Entry ID).
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

            PropertyTagArray_r propTags1 = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags1;

            // Set value for propNameOfGetMatches.
            PropertyName_r? propNameOfGetMatches = null;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            #endregion

            #region Call NspiResortRestriction method to resort MIDs got from NspiGetMatches method.
            uint reservedOfResortRestriction = 0;
            PropertyTagArray_r inmids = outMIds.Value;

            PropertyTagArray_r? outMIdOfResort = null;
            this.Result = this.ProtocolAdatper.NspiResortRestriction(reservedOfResortRestriction, ref stat, inmids, ref outMIdOfResort);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiResortRestriction should return Success!");
            #endregion

            #region Call NspiQueryRows method using MIDs of NspiResortRestriction as the input parameter lpETable to check whether the specified row is inserted.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            uint tableCount = outMIds.Value.CValues;
            uint[] table = new uint[outMIds.Value.CValues];
            Array.Copy(outMIds.Value.AulPropTag, table, outMIds.Value.CValues);
            uint count = tableCount + 1;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    (uint)AulProp.PidTagEntryId
                }
            };
            PropertyTagArray_r? propTags = propTagsInstance;

            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1198");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1198
            // If the number of the queried row equals the number of minimal entry id in outMids, 
            // it means the when each object located, a row is inserted into the constructed Explicit Table.
            Site.CaptureRequirementIfAreEqual<uint>(
                outMIdOfResort.Value.CValues,
                rowsOfQueryRows.Value.CRows,
                1198,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] For each such object located, a row is inserted into the constructed Explicit Table.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to the server ignoring some fields.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC24_ResortRestrictionIgnoreSomeFields()
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

            #region Call NspiResortRestriction method by specifying the inMIDs parameter with a Minimal Entry ID which does not specify an object.
            uint reservedOfResortRestriction = 0;
            PropertyTagArray_r inmids = new PropertyTagArray_r
            {
                CValues = 1,
                AulPropTag = new uint[1]
                {
                    uint.Parse(Constants.UnrecognizedMID)
                }
            };

            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName;
            PropertyTagArray_r? outMIds = null;
            this.Result = this.ProtocolAdatper.NspiResortRestriction(reservedOfResortRestriction, ref stat, inmids, ref outMIds);

            #region Capture
            // pStat is defined to the type of STAT in this test suite, and the CurrentRec describes the logical position. So this requirement can be captured if code can reach here.
            this.Site.CaptureRequirement(
                1172,
                @"[In NspiResortRestriction] pStat: A reference to a STAT block describing a logical position in a specific address book container.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R494");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R494
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                494,
                @"[In String Sorting] Every server MUST support sorting on Unicode string representations for the PidTagDisplayName property.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1197");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1197
            Site.CaptureRequirementIfAreEqual<uint>(
                0x0,
                outMIds.Value.CValues,
                1197,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] The server MUST ignore any Minimal Entry IDs that do not specify an object.");

            #endregion Capture
            #endregion

            #region Call NspiResortRestriction method with Reserved field set to another value different from that in the above step.
            reservedOfResortRestriction = 1;
            PropertyTagArray_r? outMIds1 = null;
            this.Result = this.ProtocolAdatper.NspiResortRestriction(reservedOfResortRestriction, ref stat, inmids, ref outMIds1);

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1915");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1915
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyTagArrayEqual(outMIds, outMIds1),
                1915,
                @"[In NspiResortRestriction] If this field [Reserved] is set to different values, the server will return the same result.");
            #endregion
            #endregion

            #region Call NspiResortRestriction method with ppOutMIds set to another value different from that in the above step.
            PropertyTagArray_r? outMIds2 = inmids;
            this.Result = this.ProtocolAdatper.NspiResortRestriction(reservedOfResortRestriction, ref stat, inmids, ref outMIds2);

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1870");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1870
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyTagArrayEqual(outMIds, outMIds2),
                1870,
                @"[In NspiResortRestriction] [ppOutMIds] If this field is set to different values, the server will return the same result.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify if the object specified by the CurrentRec field of the input parameter STAT block is not in the constructed 
        /// Explicit Table, the CurrentRec field of the output parameter STAT block is set to the value MID_BEGINNING_OF_TABLE and the NumPos field of the output 
        /// parameter STAT block is set to the value 0.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC25_ResortRestrictionCurrentRecFieldNotInETable()
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

            #region Call NspiDNToMId to get a set of Minimal Entry ID. These IDs will be used as the parameter of NspiResortRestriction.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                CValues = 3,
                LppszA = new string[3]
                {
                    Common.GetConfigurationPropertyValue("User1Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User3Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("User2Essdn", this.Site),
                }
            };

            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiDNToMId should return Success!");
            #endregion

            #region Call NspiResortRestriction with CurrentRec field not in the constructed Explicit Table.
            uint reservedOfResortRestriction = 0;
            PropertyTagArray_r inmids = new PropertyTagArray_r
            {
                CValues = mids.Value.CValues - 1
            };
            inmids.AulPropTag = new uint[inmids.CValues];
            for (int i = 0; i < inmids.CValues; i++)
            {
                inmids.AulPropTag[i] = mids.Value.AulPropTag[i];
            }

            // If the object specified by the CurrentRec field of the input parameter pStat is in the constructed Explicit Table, the NumPos field of the output parameter pStat is set to the numeric position in the Explicit Table.
            stat.CodePage = (uint)RequiredCodePage.CP_TELETEX;
            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName;
            stat.CurrentRec = mids.Value.AulPropTag[mids.Value.CValues - 1];
            PropertyTagArray_r? outMIds = null;
            this.Result = this.ProtocolAdatper.NspiResortRestriction(reservedOfResortRestriction, ref stat, inmids, ref outMIds);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiResortRestriction should return Success!");

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1205");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1205
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                stat.NumPos,
                1205,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 8 The Server MUST update the output parameter pStat as follows:] If the object specified by the CurrentRec field of the input parameter pStat is not in the constructed Explicit Table, the NumPos field of the output parameter pStat is set to the value 0.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1759");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1759
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE,
                stat.CurrentRec,
                1759,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 8 The Server MUST update the output parameter pStat as follows:] If the object specified by the CurrentRec field of the input parameter pStat is not in the constructed Explicit Table, the CurrentRec field of the output parameter pStat is set to the value MID_BEGINNING_OF_TABLE.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R80");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R80
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE,
                stat.CurrentRec,
                80,
                @"[In Positioning Minimal Entry IDs] MID_BEGINNING_OF_TABLE (0x00000000): Specifies the position before the first row in the current address book container.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiCompareMIds operations with different MID1 and MID2.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC26_CompareMIdsSuccessWithCompareMid1AndMid2()
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

            #region Call NspiDNToMId method to get two MIDs.
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

            #region Call NspiCompareMIds to compare two MIDs.
            uint firstMId = mids.Value.AulPropTag[0];
            uint secondMId = mids.Value.AulPropTag[1];

            // MId1 before MId2.
            uint reservedOfCompareMIds = 0;
            uint mid1 = firstMId;
            uint mid2 = secondMId;
            int compareResult;

            this.Result = this.ProtocolAdatper.NspiCompareMIds(reservedOfCompareMIds, stat, mid1, mid2, out compareResult);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiCompareMIds method should return Success.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1236: the returned compare result is {0}.", compareResult);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1236
            Site.CaptureRequirementIfIsTrue(
                compareResult < 0,
                1236,
                @"[In NspiCompareMIds] [Server Processing Rules: Upon receiving message NspiCompareMIds, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] If the position of the object specified by MId1 comes before the position of the object specified by MId2 in the table specified by the field ContainerID of the input parameter pStat, the server MUST return a value less than 0 in the output parameter plResult.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1242");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1242
            Site.CaptureRequirementIfAreEqual(
                ErrorCodeValue.Success,
                this.Result,
                1242,
                @"[In NspiCompareMIds] [Server Processing Rules: Upon receiving message NspiCompareMIds, the server MUST process the data from the message subject to the following constraints:] [Constraint 9] If no other return values have been specified by these constraints [constraints 1-8], the server MUST return the return value ""Success"".");
            #endregion Capture

            #endregion

            #region Swap the two MIDs and then call NspiCompareMIds to compare the two MIDs again.
            mid1 = secondMId;
            mid2 = firstMId;
            this.Result = this.ProtocolAdatper.NspiCompareMIds(reservedOfCompareMIds, stat, mid1, mid2, out compareResult);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiCompareMIds should return Success.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1238: the returned compare result is {0}.", compareResult);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1238
            Site.CaptureRequirementIfIsTrue(
                compareResult > 0,
                1238,
                @"[In NspiCompareMIds] [Server Processing Rules: Upon receiving message NspiCompareMIds, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] If the position of the object specified by MId1 comes after the position of the object specified by MId2 in the table specified by the field ContainerID of the input parameter pStat, the server MUST return a value greater than 0 in the output parameter plResult.");

            #endregion Capture
            #endregion

            #region Assign MID1 and MID2 to the same MID and then call NspiCompareMIds to compare MID1 and MID2.
            mid2 = mid1;
            this.Result = this.ProtocolAdatper.NspiCompareMIds(reservedOfCompareMIds, stat, mid1, mid2, out compareResult);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiCompareMIds should return Success.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1240");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1240
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                compareResult,
                1240,
                @"[In NspiCompareMIds] [Server Processing Rules: Upon receiving message NspiCompareMIds, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] If the position of the object specified by MId1 is the same as the position of the object specified by MId2 in the table specified by the ContainerID field of the input parameter pStat (that is, they specify the same object), the server MUST return a value of 0 in the output parameter plResult.");

            // Since the comparing work has been done in all above steps, so MS-OXNSPI_R1210 can be captured directly.
            this.Site.CaptureRequirement(
                1210,
                @"[In NspiCompareMIds] The NspiCompareMIds method compares the position in an address book container of two objects identified by Minimal Entry ID and returns the value of the comparison.");
            #endregion

            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify in NspiCompareMIds if parameter Reserved is set to different values, the server will return the same result.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC27_CompareMIdsIgnoreReserved()
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

            #region Call NspiDNToMId to get a set of Minimal Entry ID. These IDs will be used as the parameter of NspiCompareMIds.
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

            #region Call NspiCompareMIds with Reserved set to 0.
            uint reservedOfCompareMIds = 0;
            uint mid1 = mids.Value.AulPropTag[0];
            uint mid2 = mids.Value.AulPropTag[1];
            int results1;
            this.Result = this.ProtocolAdatper.NspiCompareMIds(reservedOfCompareMIds, stat, mid1, mid2, out results1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiCompareMIds should return Success!");
            #endregion

            #region Call NspiCompareMIds with Reserved set to 1.
            int results2;
            reservedOfCompareMIds = 0x1;
            this.Result = this.ProtocolAdatper.NspiCompareMIds(reservedOfCompareMIds, stat, mid1, mid2, out results2);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiCompareMIds should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1699");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1699
            Site.CaptureRequirementIfAreEqual<int>(
                results1,
                results2,
                1699,
                @"[In NspiCompareMIds] If this field[Reserved] is set to different values, the server will return the same result.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiDNToMId operation returning success.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC28_DNToMIdSuccess()
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

            #region Call NspiQueryRows method to get a set of valid rows.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fSkipObjects;
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
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");

            // Get mid and entry id of corresponding name.
            string administratorName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            string mailUserName = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            // The DN string of administrator.
            string administratorESSDN = string.Empty;

            // The DN string of mailUserName.
            string mailUserNameESSDN = string.Empty;

            // Parse PermanentEntryID in rows to get the DN of administrator and mailUserName.
            for (int i = 0; i < rowsOfQueryRows.Value.CRows; i++)
            {
                string name = System.Text.Encoding.UTF8.GetString(rowsOfQueryRows.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(administratorName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    PermanentEntryID administratorEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.Bin.Lpb);
                    administratorESSDN = administratorEntryID.DistinguishedName;
                }
                else if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(mailUserName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    PermanentEntryID mailUserEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.Bin.Lpb);
                    mailUserNameESSDN = mailUserEntryID.DistinguishedName;
                }

                if (!string.IsNullOrEmpty(administratorESSDN) && !string.IsNullOrEmpty(mailUserNameESSDN))
                {
                    break;
                }
            }
            #endregion

            #region Call NspiDNToMId method to map the DNs to a set of Minimal Entry ID.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                CValues = 2,
                LppszA = new string[2]
            };
            names.LppszA[0] = administratorESSDN;
            names.LppszA[1] = mailUserNameESSDN;
            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1266");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1266
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1266,
                @"[In NspiDNToMId] [Server Processing Rules: Upon receiving message NspiDNToMId, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If no other return values have been specified by these constraints [constraints 1-4], the server MUST return the return value ""Success"".");

            #endregion Capture
            #endregion

            #region Call NspiUpdateStat to update the positioning changes in a table.
            reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiGetProps to get the Permanent Entry ID using one Minimal Entry ID returned from step 3.
            PropertyTagArray_r prop = new PropertyTagArray_r
            {
                CValues = 1
            };
            prop.AulPropTag = new uint[prop.CValues];
            prop.AulPropTag[0] = (uint)AulProp.PidTagEntryId;

            propTags = prop;
            PropertyRow_r? rows;

            uint flagsOfGetProps = (uint)RetrievePropertyFlag.fSkipObjects;
            stat.CurrentRec = mids.Value.AulPropTag[0];

            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetProps should return Success.");
            PermanentEntryID entryID1 = AdapterHelper.ParsePermanentEntryIDFromBytes(rows.Value.LpProps[0].Value.Bin.Lpb);
            bool firstDNMapright = entryID1.DistinguishedName.Equals(names.LppszA[0], StringComparison.OrdinalIgnoreCase);
            #endregion

            #region Call NspiGetProps to get the Permanent Entry ID using another Minimal Entry ID returned from step 3.
            stat.CurrentRec = mids.Value.AulPropTag[1];
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetProps should return Success.");
            PermanentEntryID entryID2 = AdapterHelper.ParsePermanentEntryIDFromBytes(rows.Value.LpProps[0].Value.Bin.Lpb);
            bool secondDNMapright = entryID2.DistinguishedName.Equals(names.LppszA[1], StringComparison.OrdinalIgnoreCase);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1263: the first DN map right result is {0}, the second DN map right result is {1}", firstDNMapright, secondDNMapright);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1263
            Site.CaptureRequirementIfIsTrue(
                firstDNMapright && secondDNMapright,
                1263,
                @"[In NspiDNToMId] [Server Processing Rules: Upon receiving message NspiDNToMId, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] The list is in a one-to-one order preserving correspondence with the list of DNs in the input parameter pNames.");

            // Since the mapping work has been done in this step, so MS-OXNSPI_R1243 can be captured directly.
            this.Site.CaptureRequirement(
                1243,
                @"[In NspiDNToMId] The NspiDNToMId method maps a set of DNs to a set of Minimal Entry ID.");
            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiDNToMId operation with a DN that is unable to be located.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC29_DNToMIdUnableToLocate()
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

            #region Call NspiDNToMId method with a DN which is unable to be located.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                CValues = 1,
                LppszA = new string[1]
            };
            names.LppszA[0] = "UnableToLocate";
            PropertyTagArray_r? mids;

            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiDNToMId method should return Success.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1260");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1260
            Site.CaptureRequirementIfAreEqual<uint>(
                0x0,
                mids.Value.AulPropTag[0],
                1260,
                @"[In NspiDNToMId] [Server Processing Rules: Upon receiving message NspiDNToMId, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the server is unable to locate an appropriate mapping between a DN and a Minimal Entry ID, it [server] MUST map the DN to a Minimal Entry ID with the value 0.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify in NspiDNToMId if the parameter Reserved is set to different values, the server will return the same result. 
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC30_DNToMIdIgnoreReserved()
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

            #region Call NspiDNToMId with Reserved field set to 0.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                CValues = 2,
                LppszA = new string[2]
            };
            names.LppszA[0] = Common.GetConfigurationPropertyValue("User3Essdn", this.Site);
            names.LppszA[1] = Common.GetConfigurationPropertyValue("User2Essdn", this.Site);
            PropertyTagArray_r? mids1;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiDNToMId should return Success!");
            #endregion

            #region Call NspiDNToMId with Reserved field set to 1.
            PropertyTagArray_r? mids2;
            reserved = 0x1;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids2);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiDNToMId method should return Success.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1700");

            bool temp = true;
            if (mids1.Value.CValues == mids2.Value.CValues)
            {
                for (int i = 0; i < mids1.Value.CValues; i++)
                {
                    if (mids1.Value.AulPropTag[i] != mids2.Value.AulPropTag[i])
                    {
                        temp = false;
                        break;
                    }
                }
            }
            else
            {
                temp = false;
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1700
            Site.CaptureRequirementIfIsTrue(
                temp,
                1700,
                @"[In NspiDNToMId] if this field[Reserved] is set to different values, the server will return the same result.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetMatches operation with different restrictions.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC31_NspiGetMatchesRestrictionVerification()
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

            #region Call NspiGetMatches with ExistRestriction and AndRestriction.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;

            Restriction_r resExist_r = new Restriction_r
            {
                Rt = 0x08,
                Res = new RestrictionUnion_r
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

            Restriction_r resExist_r1 = new Restriction_r
            {
                Rt = 0x08,
                Res =
                    new RestrictionUnion_r
                    {
                        ResExist =
                            new ExistRestriction_r
                            {
                                Reserved1 = 0,
                                Reserved2 = 0,
                                PropTag = (uint)AulProp.PidTagAddressBookPhoneticDisplayName
                            }
                    }
            };

            Restriction_r addRestriction = new Restriction_r
            {
                Rt = 0x00,
                Res =
                    new RestrictionUnion_r
                    {
                        ResAnd = new AndRestriction_r
                        {
                            CRes = 2,
                            LpRes = new Restriction_r[]
                            {
                                resExist_r, resExist_r1
                            }
                        }
                    }
            };

            Restriction_r? filter = addRestriction;

            PropertyTagArray_r propTags1 = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            PropertyTagArray_r? propTagsOfGetMatches = propTags1;

            // Set value for propNameOfGetMatches.
            PropertyName_r? propNameOfGetMatches = null;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            uint sortType = stat.SortType;
            uint currentRec = stat.CurrentRec;
            int delta = stat.Delta;
            uint numPos = stat.NumPos;
            uint totalRecs = stat.TotalRecs;
            uint codePage = stat.CodePage;
            uint templateLocale = stat.TemplateLocale;
            uint sortLocale = stat.SortLocale;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            int index = 0;
            foreach (PropertyRow_r rows in rowsOfGetMatches.Value.ARow)
            {
                uint propTag = outMIds.Value.AulPropTag[index];

                // Fetch the value of PidTagEntryId.
                EphemeralEntryID entryId = AdapterHelper.ParseEphemeralEntryIDFromBytes(rows.LpProps[0].Value.Bin.Lpb);
                Site.Assert.AreEqual<uint>(propTag, entryId.Mid, "The Minimal ID of the object (index {0}) is inserted into the Explicit Table.", index);
                index++;
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1144
            // The Minimal ID of the object is inserted into the Explicit Table according to the above assert, so MS-OXNSPI_R1144 can be verified directly.
            Site.CaptureRequirement(
                1144,
                "[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 13] If a specific object is located, the Minimal ID of the object is inserted into the Explicit Table.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1146");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1146
            Site.CaptureRequirementIfAreEqual<uint>(
                currentRec,
                stat.ContainerID,
                1146,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 15] If the server returns ""Success"", the server MUST set the ContainerID field of the output parameter pStat to be equal to the CurrentRec field of the input parameter pStat.");

            Site.Assert.AreEqual<uint>(sortType, stat.SortType, "The server MUST NOT modify SortType field in pStat parameter.");
            Site.Assert.AreEqual<uint>(currentRec, stat.CurrentRec, "The server MUST NOT modify CurrentRec field in pStat parameter.");
            Site.Assert.AreEqual<int>(delta, stat.Delta, "The server MUST NOT modify Delta field in pStat parameter.");
            Site.Assert.AreEqual<uint>(numPos, stat.NumPos, "The server MUST NOT modify NumPos field in pStat parameter.");
            Site.Assert.AreEqual<uint>(totalRecs, stat.TotalRecs, "The server MUST NOT modify TotalRecs field in pStat parameter.");
            Site.Assert.AreEqual<uint>(codePage, stat.CodePage, "The server MUST NOT modify CodePage field in pStat parameter.");
            Site.Assert.AreEqual<uint>(templateLocale, stat.TemplateLocale, "The server MUST NOT modify TemplateLocale field in pStat parameter.");
            Site.Assert.AreEqual<uint>(sortLocale, stat.SortLocale, "The server MUST NOT modify SortLocale field in pStat parameter.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1147
            // The server didn't modify any other fields except ContainerID field in pStat parameter, MS-OXNSPI_R1147 can be verified directly.
            Site.CaptureRequirement(
                1147,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 15] The server MUST NOT modify any other fields in this parameter [pStat].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1908");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1908
            // The AndRestriction_r is specified to the server and the NspiGetMatches method performs successfully, if the list of Minimal Entry IDs is not null, MS-OXNSPI_R1908 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                outMIds,
                1908,
                @"[In NspiGetMatches] When AndRestriction_r is specified to the server via the NspiGetMatches method, the server locates all the objects that meet the restriction criteria, and the list of the Minimal Entry IDs of those objects is constructed.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1914");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1914
            // The ExistRestriction_r is specified to the server and the NspiGetMatches method performs successfully, if the list of Minimal Entry IDs is not null, MS-OXNSPI_R1914 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                outMIds,
                1914,
                @"[In NspiGetMatches] When ExistRestriction_r is specified to the server via the NspiGetMatches method, the server locates all the objects that meet the restriction criteria, and the list of the Minimal Entry IDs of those objects is constructed.");
            #endregion

            #region Call NspiGetMatches with ExistRestriction and OrRestriction.

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
                                resExist_r, resExist_r1
                            }
                        }
                    }
            };

            filter = restrictionOr;

            propTags1 = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            propTagsOfGetMatches = propTags1;

            // Set value for propNameOfGetMatches.
            propNameOfGetMatches = null;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1908");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1909
            // The OrRestriction_r is specified to the server and the NspiGetMatches method performs successfully, if the list of Minimal Entry IDs is not null, MS-OXNSPI_R1909 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                outMIds,
                1909,
                @"[In NspiGetMatches] When OrRestriction_r is specified to the server via the NspiGetMatches method, the server locates all the objects that meet the restriction criteria, and the list of the Minimal Entry IDs of those objects is constructed.");
            #endregion

            #region Call NspiGetMatches with ContentRestriction.
            Restriction_r contentRestriction = new Restriction_r
            {
                Rt = 0x03,
                Res = new RestrictionUnion_r
                {
                    ResContent = new ContentRestriction_r
                    {
                        FuzzyLevel = 0x00010002 // The value stored in the TaggedValue field matches a starting portion of the value of the column property tag and the comparison does not consider case.
                    }
                }
            };

            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = 0
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

            contentRestriction.Res.ResContent.Prop = new PropertyValue_r[] { target };
            contentRestriction.Res.ResContent.PropTag = (uint)AulProp.PidTagDisplayName;

            filter = contentRestriction;

            propTags1 = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            propTagsOfGetMatches = propTags1;

            // Set value for propNameOfGetMatches.
            propNameOfGetMatches = null;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1911");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1911
            // The ContentRestriction_r is specified to the server and the NspiGetMatches method performs successfully, if the list of Minimal Entry IDs is not null, MS-OXNSPI_R1911 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                outMIds,
                1911,
                @"[In NspiGetMatches] When ContentRestriction_r is specified to the server via the NspiGetMatches method, the server locates all the objects that meet the restriction criteria, and the list of the Minimal Entry IDs of those objects is constructed.");

            #endregion

            #region Call NspiGetMatches with PropertyRestriction.
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
            target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = 0
            };
            displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);
            }
            else
            {
                target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName + "\0");
            }

            propertyRestriction.Res.ResProperty.Prop = new PropertyValue_r[] { target };
            propertyRestriction.Res.ResProperty.PropTag = (uint)AulProp.PidTagDisplayName;

            filter = propertyRestriction;

            propTags1 = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };
            propTagsOfGetMatches = propTags1;

            // Set null value to input parameter property names of NspiGetMatches.
            propNameOfGetMatches = null;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1912");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1912
            // The PropertyRestriction_r is specified to the server and the NspiGetMatches method performs successfully, if the list of Minimal Entry IDs is not null, MS-OXNSPI_R1912 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                outMIds,
                1912,
                @"[In NspiGetMatches] When PropertyRestriction_r is specified to the server via the NspiGetMatches method, the server locates all the objects that meet the restriction criteria, and the list of the Minimal Entry IDs of those objects is constructed.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1913
            // AndRestriction_r, OrRestriction_r, ContentRestriction_r, PropertyRestriction_r and ExistRestriction_r are verified successfully, MS-OXNSPI_R1913 can be verified directly.
            Site.CaptureRequirement(
                1913,
                @"[In NspiGetMatches] When RestrictionUnion_r is specified to the server via the NspiGetMatches method, the server locates all the objects that meet the restriction criteria, and the list of the Minimal Entry IDs of those objects is constructed.");
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to the string values returned from NspiGetProps operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC32_NspiGetPropsStringConversion()
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

            #region Call NspiUpdateStat to update the STAT block.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiGetProps to request the string properties to be different from their native types.
            PropertyTagArray_r prop = new PropertyTagArray_r
            {
                CValues = 2
            };
            prop.AulPropTag = new uint[prop.CValues];

            prop.AulPropTag[0] = AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable);
            prop.AulPropTag[1] = AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName);
            PropertyTagArray_r? propTags = prop;

            uint flagsOfGetProps = (uint)RetrievePropertyFlag.fEphID;
            PropertyRow_r? rows;

            // Save the CurrentRec value of stat structure.
            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetProps should return Success.");
            Site.Assert.IsNotNull(rows, "The reference to the PropertyRow_r value should not be null. The row number is {0}.", rows == null ? 0 : rows.Value.CValues);

            string diplayName = System.Text.UnicodeEncoding.Unicode.GetString(rows.Value.LpProps[1].Value.LpszW);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1968: The value of PidTagDisplayNamePrintable of the specified address book returned in Unicode representation is {0}.", diplayName);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1968
            // The field LpszW indicates a single Unicode string value.
            Site.CaptureRequirementIfIsNotNull(
                rows.Value.LpProps[1].Value.LpszW,
                1968,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiGetProps method, String values can be returned in Unicode representation in the output parameter ppRows.");

            string displayNamePrintable = System.Text.UnicodeEncoding.UTF8.GetString(rows.Value.LpProps[0].Value.LpszA);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1969: The value of PidTagAddressBookDisplayNamePrintable of the specified address book returned in Unicode representation is {0}.", displayNamePrintable);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1969
            // The field LpszA indicates a 8-bit character string value.
            Site.CaptureRequirementIfIsNotNull(
                rows.Value.LpProps[0].Value.LpszA,
                1969,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiGetProps method, String values can be returned in 8-bit character representation in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1934");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1934
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                rows.Value.LpProps[0].PropTag,
                1934,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetProps] If the native type of a property is PtypString and the client has requested that property with the type PtypString8, the server MUST convert the Unicode representation to an 8-bit character representation in the code page specified by the CodePage field of the pStat parameter prior to returning the value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1950");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1950
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName),
                rows.Value.LpProps[1].PropTag,
                1950,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetProps] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString, the server MUST convert the 8-bit character representation to a Unicode representation prior to returning the value.");
            #endregion

            #region Call NspiGetProps to request the string properties to be the same as their native types.
            prop = new PropertyTagArray_r
            {
                CValues = 2
            };
            prop.AulPropTag = new uint[prop.CValues];

            prop.AulPropTag[0] = (uint)AulProp.PidTagAddressBookDisplayNamePrintable;
            prop.AulPropTag[1] = (uint)AulProp.PidTagDisplayName;
            propTags = prop;

            this.Result = this.ProtocolAdatper.NspiGetProps(flagsOfGetProps, stat, propTags, out rows);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetProps should return Success.");
            Site.Assert.IsNotNull(rows, "The reference to the PropertyRow_r value should not be null. The row number is {0}.", rows == null ? 0 : rows.Value.CValues);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1942");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1942
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                rows.Value.LpProps[0].PropTag,
                1942,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetProps] If the native type of a property is PtypString and the client has requested that property with the type PtypString, the server MUST return the Unicode representation unmodified.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1958");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1958
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagDisplayName,
                rows.Value.LpProps[1].PropTag,
                1958,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetProps] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString8, the server MUST return the 8-bit character representation unmodified.");

            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to string values returned from NspiQueryRows operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC33_NspiQueryRowsStringConversion()
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

            #region Call NspiQueryRows to request the string properties to be different from their native types.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fEphID;
            PropertyRowSet_r? rowsOfQueryRows;

            uint tableCount = 0;
            uint[] table = null;
            uint count = 10;
            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r
            {
                AulPropTag = new uint[]
                {
                    AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                    AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName)
                }
            };
            propTagsInstance.CValues = (uint)propTagsInstance.AulPropTag.Length;
            PropertyTagArray_r? propTags = propTagsInstance;

            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");
            Site.Assert.IsNotNull(rowsOfQueryRows, "PropertyRowSet_r value should not null. The row number is {0}.", rowsOfQueryRows == null ? 0 : rowsOfQueryRows.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowsOfQueryRows.Value.ARow.Length, "At least one address book object should be matched.");

            PropertyRow_r rowValue = rowsOfQueryRows.Value.ARow[0];

            string displayName = System.Text.UnicodeEncoding.Unicode.GetString(rowValue.LpProps[1].Value.LpszW);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1970: The value of PidTagDisplayName is {0}.", displayName);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1970
            // The field LpszW indicates a single Unicode string value.
            Site.CaptureRequirementIfIsNotNull(
                rowValue.LpProps[1].Value.LpszW,
                1970,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiQueryRows method, String values can be returned in Unicode representation in the output parameter ppRows.");

            string displayNamePrintable = System.Text.UnicodeEncoding.UTF8.GetString(rowValue.LpProps[0].Value.LpszA);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1971: The value of PidTagAddressBookDisplayNamePrintable is {0}.", displayNamePrintable);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1971
            // The field LpszA indicates a 8-bit character string value.
            Site.CaptureRequirementIfIsNotNull(
                rowValue.LpProps[0].Value.LpszA,
                1971,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiQueryRows method, String values can be returned in 8-bit character representation in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1935");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1935
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                rowValue.LpProps[0].PropTag,
                1935,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiQueryRows] If the native type of a property is PtypString and the client has requested that property with the type PtypString8, the server MUST convert the Unicode representation to an 8-bit character representation in the code page specified by the CodePage field of the pStat parameter prior to returning the value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1951");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1951
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName),
                rowValue.LpProps[1].PropTag,
                1951,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiQueryRows] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString, the server MUST convert the 8-bit character representation to a Unicode representation prior to returning the value.");
            #endregion

            #region Call NspiQueryRows to request the string properties to be the same as their native types.
            propTagsInstance = new PropertyTagArray_r
            {
                AulPropTag = new uint[]
                {
                    (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                    (uint)AulProp.PidTagDisplayName
                }
            };
            propTagsInstance.CValues = (uint)propTagsInstance.AulPropTag.Length;
            propTags = propTagsInstance;

            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");
            Site.Assert.IsNotNull(rowsOfQueryRows, "PropertyRowSet_r value should not null. The row number is {0}.", rowsOfQueryRows == null ? 0 : rowsOfQueryRows.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowsOfQueryRows.Value.ARow.Length, "At least one address book object should be matched.");

            rowValue = rowsOfQueryRows.Value.ARow[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1943");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1943
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                rowValue.LpProps[0].PropTag,
                1943,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiQueryRows] If the native type of a property is PtypString and the client has requested that property with the type PtypString, the server MUST return the Unicode representation unmodified.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1959");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1959
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagDisplayName,
                rowValue.LpProps[1].PropTag,
                1959,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiQueryRows] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString8, the server MUST return the 8-bit character representation unmodified.");

            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to string values returned from NspiGetMatches operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC34_NspiGetMatchesStringConversion()
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

            #region Call NspiGetMatches to request the string properties to be different from their native types.
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
                AulPropTag = new uint[]
                {
                    AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                    AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName)
                }
            };
            propTags.CValues = (uint)propTags.AulPropTag.Length;
            PropertyTagArray_r? propTagsOfGetMatches = propTags;
            PropertyName_r? propNameOfGetMatches = null;

            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            stat.SortType = (uint)TableSortOrder.SortTypeDisplayName_RO;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            Site.Assert.IsNotNull(rowsOfGetMatches, "PropertyRowSet_r value should not null. The row number is {0}.", rowsOfGetMatches == null ? 0 : rowsOfGetMatches.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowsOfGetMatches.Value.ARow.Length, "At least one address book object should be matched.");

            PropertyRow_r rowValue = rowsOfGetMatches.Value.ARow[0];

            string displayName = System.Text.UnicodeEncoding.Unicode.GetString(rowValue.LpProps[1].Value.LpszW);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1972: The value of PidTagDisplayName is {0}.", displayName);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1972
            // The field LpszW indicates a single Unicode string value.
            Site.CaptureRequirementIfIsNotNull(
                rowValue.LpProps[1].Value.LpszW,
                1972,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiGetMatches method, String values can be returned in Unicode representation in the output parameter ppRows.");

            string displayNamePrintable = System.Text.UnicodeEncoding.UTF8.GetString(rowValue.LpProps[0].Value.LpszA);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1973: The value of PidTagAddressBookDisplayNamePrintable is {0}.", displayNamePrintable);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1973
            // The field LpszA indicates a 8-bit character string value.
            Site.CaptureRequirementIfIsNotNull(
                rowValue.LpProps[0].Value.LpszA,
                1973,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiGetMatches method, String values can be returned in 8-bit character representation in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1936");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1936
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                rowValue.LpProps[0].PropTag,
                1936,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetMatches] If the native type of a property is PtypString and the client has requested that property with the type PtypString8, the server MUST convert the Unicode representation to an 8-bit character representation in the code page specified by the CodePage field of the pStat parameter prior to returning the value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1952");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1952
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName),
                rowValue.LpProps[1].PropTag,
                1952,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetMatches] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString, the server MUST convert the 8-bit character representation to a Unicode representation prior to returning the value.");
            #endregion

            #region Call NspiGetMatches to request the string properties to be the same as their native types.
            propTags = new PropertyTagArray_r
            {
                AulPropTag = new uint[]
                {
                    (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                    (uint)AulProp.PidTagDisplayName
                }
            };
            propTags.CValues = (uint)propTags.AulPropTag.Length;
            propTagsOfGetMatches = propTags;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            Site.Assert.IsNotNull(rowsOfGetMatches, "PropertyRowSet_r value should not null. The row number is {0}.", rowsOfGetMatches == null ? 0 : rowsOfGetMatches.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowsOfGetMatches.Value.ARow.Length, "At least one address book object should be matched.");

            rowValue = rowsOfGetMatches.Value.ARow[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1944");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1944
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                rowValue.LpProps[0].PropTag,
                1944,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetMatches] If the native type of a property is PtypString and the client has requested that property with the type PtypString, the server MUST return the Unicode representation unmodified.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1960");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1960
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagDisplayName,
                rowValue.LpProps[1].PropTag,
                1960,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetMatches] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString8, the server MUST return the 8-bit character representation unmodified.");
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to string values returned from NspiSeekEntries operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC35_NspiSeekEntriesStringConversion()
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

            #region Call NspiUpdateStat to update the STAT block that represents the position in a table to reflect positioning changes requested by the client.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiGetPropList method with dwFlags set to fEphID and CodePage field of STAT block not set to CP_WINUNICODE.
            uint flagsOfGetPropList = (uint)RetrievePropertyFlag.fEphID;
            PropertyTagArray_r? propTagsOfGetPropList;
            uint codePage = (uint)RequiredCodePage.CP_TELETEX;

            this.Result = this.ProtocolAdatper.NspiGetPropList(flagsOfGetPropList, stat.CurrentRec, codePage, out propTagsOfGetPropList);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetPropList should return Success!");
            #endregion

            #region Call NspiSeekentries to request the string properties to be different with their native types.
            uint reservedOfSeekEntries = 0;

            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName),
                Reserved = (uint)0x00,
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

            target.Value.LpszW = System.Text.Encoding.Unicode.GetBytes(displayName);

            PropertyTagArray_r propTags = new PropertyTagArray_r
            {
                AulPropTag = new uint[]
                {
                    AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                    AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName)
                }
            };
            propTags.CValues = (uint)propTags.AulPropTag.Length;
            PropertyTagArray_r? propTagsOfSeekEntries = propTags;

            PropertyTagArray_r? tableOfSeekEntries = null;
            PropertyRowSet_r? rowsOfSeekEntries;

            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiSeekEntries should return Success!");
            Site.Assert.IsNotNull(rowsOfSeekEntries, "PropertyRowSet_r value should not null. The row number is {0}.", rowsOfSeekEntries == null ? 0 : rowsOfSeekEntries.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowsOfSeekEntries.Value.ARow.Length, "At least one address book object should be matched.");
            PropertyRow_r rowValue = rowsOfSeekEntries.Value.ARow[0];

            displayName = System.Text.UnicodeEncoding.Unicode.GetString(rowValue.LpProps[1].Value.LpszW);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1974: The value of PidTagDisplayName is {0}.", displayName);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1974
            // The field LpszW indicates a single Unicode string value.
            Site.CaptureRequirementIfIsNotNull(
                rowValue.LpProps[1].Value.LpszW,
                1974,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiSeekEntries method, String values can be returned in Unicode representation in the output parameter ppRows.");

            string displayNamePrintable = System.Text.UnicodeEncoding.UTF8.GetString(rowValue.LpProps[0].Value.LpszA);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1975: The value of PidTagAddressBookDisplayNamePrintable is {0}.", displayNamePrintable);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1975
            // The field LpszA indicates a 8-bit character string value.
            Site.CaptureRequirementIfIsNotNull(
                rowValue.LpProps[0].Value.LpszA,
                1975,
                "[In Conversion Rules for String Values Specified by the Server to the Client] In NspiSeekEntries method, String values can be returned in 8-bit character representation in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1937");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1937
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertStringToString8((uint)AulProp.PidTagAddressBookDisplayNamePrintable),
                rowValue.LpProps[0].PropTag,
                1937,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiSeekEntries] If the native type of a property is PtypString and the client has requested that property with the type PtypString8, the server MUST convert the Unicode representation to an 8-bit character representation in the code page specified by the CodePage field of the pStat parameter prior to returning the value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1953");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1953
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                AdapterHelper.ConvertString8ToString((uint)AulProp.PidTagDisplayName),
                rowValue.LpProps[1].PropTag,
                1953,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiSeekEntries] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString, the server MUST convert the 8-bit character representation to a Unicode representation prior to returning the value.");
            #endregion

            #region Call NspiSeekEntries to request the string properties to be the same as their native types.

            target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = (uint)0x00,
            };
            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            }
            else
            {
                displayName = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "\0";
            }

            target.Value.LpszA = System.Text.Encoding.UTF8.GetBytes(displayName);

            propTags = new PropertyTagArray_r
            {
                AulPropTag = new uint[]
                {
                    (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                    (uint)AulProp.PidTagDisplayName
                }
            };
            propTags.CValues = (uint)propTags.AulPropTag.Length;
            propTagsOfSeekEntries = propTags;

            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiSeekEntries should return Success!");
            Site.Assert.IsNotNull(rowsOfSeekEntries, "PropertyRowSet_r value should not null. The row number is {0}.", rowsOfSeekEntries == null ? 0 : rowsOfSeekEntries.Value.CRows);
            Site.Assert.AreNotEqual<int>(0, rowsOfSeekEntries.Value.ARow.Length, "At least one address book object should be matched.");

            rowValue = rowsOfSeekEntries.Value.ARow[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1945");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1945
            // The native type of PidTagAddressBookDisplayNamePrintable is PtypString and the client has requested this property with PtypString.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagAddressBookDisplayNamePrintable,
                rowValue.LpProps[0].PropTag,
                1945,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiSeekEntries] If the native type of a property is PtypString and the client has requested that property with the type PtypString, the server MUST return the Unicode representation unmodified.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1961");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1961
            // The native type of PidTagDisplayName is PtypString8 and the client has requested this property with PtypString8.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagDisplayName,
                rowValue.LpProps[1].PropTag,
                1961,
                "[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiSeekEntries] If the native type of a property is PtypString8 and the client has requested that property with the type PtypString8, the server MUST return the 8-bit character representation unmodified.");
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiSeekEntries with pPropTags set to a null value.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC36_NspiSeekEntriesWithpPropTagsSetNull()
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

            #region Call NspiQueryRows method with propTags set to null.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fSkipObjects;
            uint tableCount = 0;
            uint[] table = null;
            uint count = Constants.QueryRowsRequestedRowNumber;
            PropertyRowSet_r? rowsOfQueryRows;
            string displayName = Common.GetConfigurationPropertyValue("AgentName", this.Site);

            PropertyTagArray_r? propTags = null;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success.");
            Site.Assert.IsNotNull(rowsOfQueryRows.Value.ARow, "The returned rows should not be empty. The row number is {0}", rowsOfQueryRows == null ? 0 : rowsOfQueryRows.Value.CRows);

            uint[] outMids1 = new uint[rowsOfQueryRows.Value.CRows];
            uint position = 0xffffffff;
            for (uint i = 0; i < rowsOfQueryRows.Value.CRows; i++)
            {
                string name = System.Text.Encoding.UTF8.GetString(rowsOfQueryRows.Value.ARow[i].LpProps[3].Value.LpszA);
                if (displayName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                {
                    position = i;
                }

                outMids1[i] = (uint)rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.L;
            }
            #endregion

            #region Call NspiUpdateStat to update the STAT block that represents the position in a table to reflect positioning changes requested by the client.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiSeekEntries method with requesting 1 prop tag.
            uint reservedOfSeekEntries = 0;
            PropertyValue_r target = new PropertyValue_r
            {
                PropTag = (uint)AulProp.PidTagDisplayName,
                Reserved = (uint)0x00
            };
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
                    (uint)AulProp.PidTagDisplayName
                }
            };
            PropertyTagArray_r? propTagsOfSeekEntries = tags;

            PropertyTagArray_r table1 = new PropertyTagArray_r
            {
                AulPropTag = outMids1
            };
            table1.CValues = (uint)table1.AulPropTag.Length;
            PropertyTagArray_r? tableOfSeekEntries = table1;
            PropertyRowSet_r? rowsOfSeekEntries;

            stat.CurrentRec = 0;
            stat.NumPos = 0;
            this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiSeekEntries should return Success!");

            Site.CaptureRequirementIfAreEqual<uint>(
                position,
                stat.NumPos,
                1052,
                "[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] If the server is using the table specified by the input parameter lpETable, the server MUST set the NumPos field of the parameter pStat to the accurate numeric position of the qualifying row in the table.");

            #endregion

            #region Call NspiQueryRows method using parameters of NspiSeekEntries method.
            uint flagsOfQueryRows1 = (uint)RetrievePropertyFlag.fEphID;

            // Create a new table used as the NspiQueryRows input parameter.
            uint[] newTable = new uint[outMids1.Length - stat.NumPos];
            Array.Copy(outMids1, stat.NumPos, newTable, 0, newTable.Length);
            uint tableCount1 = (uint)newTable.Length;
            uint count1 = tableCount1;
            PropertyRowSet_r? rowsOfQueryRows1;
            PropertyTagArray_r? propTags1 = tags;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows1, ref stat, tableCount1, newTable, count1, propTags1, out rowsOfQueryRows1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");

            // The returned rows may be empty for the NspiSeekEntries method as indicated in section 3.1.4.1.9 of the specification MS-OXNSPI. SO if it is empty, this method needs to be called again.
            if (rowsOfSeekEntries == null || rowsOfSeekEntries.Value.ARow.Length != rowsOfQueryRows1.Value.ARow.Length)
            {
                this.Result = this.ProtocolAdatper.NspiSeekEntries(reservedOfSeekEntries, ref stat, target, tableOfSeekEntries, propTagsOfSeekEntries, out rowsOfSeekEntries);
                Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiSeekEntries should return Success!");
            }

            Site.Assert.IsNotNull(rowsOfSeekEntries, "The NspiSeekEntries method should return rows when called again!");

            #region Capture
            foreach (PropertyRow_r row in rowsOfSeekEntries.Value.ARow)
            {
                Site.Assert.AreEqual<int>(1, row.LpProps.Length, "The input parameter pPropTags specified 1 property tag.");
                Site.Assert.AreEqual<uint>((uint)AulProp.PidTagDisplayName, row.LpProps[0].PropTag, "The input parameter pPropTags specified property tag PidTagDisplayName.");
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1058
            // All the property tag returned in ppRows were checked and same as specified by the client, MS-OXNSPI_R1058 can be verified directly.
            Site.CaptureRequirement(
                1058,
                @"[In NspiSeekEntries] [Server Processing Rules: Upon receiving message NspiSeekEntries, the server MUST process the data from the message subject to the following constraints:] [Constraint 15] Subject to the prior constraints [If the input parameter pPropTags is not NULL], the server MUST construct an PropertyRowSet_r to return to the client in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1755");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1755
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowSetEqual(rowsOfSeekEntries, rowsOfQueryRows1),
                1755,
                @"This PropertyRowSet_r MUST be exactly the same PropertyRowSet_r that would be returned in the ppRows parameter of a call to the NspiQueryRows method with the following parameters:
The NspiSeekEntries parameter hRpc is used as the NspiQueryRows parameter hRpc.
The value fEphID is used as the NspiQueryRows parameter dwFlags.
The NspiSeekEntries output parameter pStat (as modified by the prior constraints) is used as the NspiQueryRows parameter pStat.
If the NspiSeekEntries input parameter lpETable is not NULL, the server constructs an explicit table from the table specified by lpETable by copying rows in order from lpETable to the new explicit table. The server begins copying from the row specified by the NumPos field of the pStat parameter (as modified by the prior constraints), and continues until all remaining rows are added to the new table. The number of rows in this new table is used as the NspiQueryRows parameter dwETableCount, and the new table is used as the NspiQueryRows lpETable parameter.
The list of Minimal Entry IDs in the input parameter lpETable starting with the qualifying row is used as the NspiQueryRows parameter lpETable. These Minimal Entry IDs [the list of Minimal Entry IDs in the input parameter lpETable] are expressed as a simple array of DWORD values rather than as a PropertyTagArray_r value. Note that the qualifying row is included in this list, and that the order of the Minimal Entry IDs from the input parameter lpETable is preserved in this list.
If the NspiSeekEntries input parameter lpETable is not NULL, the value used for the NspiQueryRows parameter dwETableCount is used for the NspiQueryRows parameter Count.
The NspiSeekEntries parameter pPropTags is used as the NspiQueryRows parameter pPropTags.");
            #endregion Capture
            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the Minimal Entry ID.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC37_MinimalEntryIDVerification()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

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

            #region Call NspiQueryRows method to get a set of valid rows.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fSkipObjects;
            uint tableCount = 0;
            uint[] table = null;
            uint count = Constants.QueryRowsRequestedRowNumber;
            PropertyRowSet_r? rowsOfQueryRows;

            PropertyTagArray_r propTagsInstance = new PropertyTagArray_r();
            propTagsInstance.CValues = 4;
            propTagsInstance.AulPropTag = new uint[4]
            {     
                (uint)AulProp.PidTagEntryId,
                (uint)AulProp.PidTagDisplayName, 
                (uint)AulProp.PidTagDisplayType, 
                (uint)AulProp.PidTagAddressBookMember
            };
            PropertyTagArray_r? propTags = propTagsInstance;
            this.Result = this.ProtocolAdatper.NspiQueryRows(flagsOfQueryRows, ref stat, tableCount, table, count, propTags, out rowsOfQueryRows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");

            // Get mid and entry id of corresponding name.
            string administratorName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            string mailUserName = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            // The DN string of administrator.
            string administratorESSDN = string.Empty;

            // The DN string of mailUserName.
            string mailUserNameESSDN = string.Empty;

            // Parse PermanentEntryID in rows to get the DN of administrator and mailUserName.
            for (int i = 0; i < rowsOfQueryRows.Value.CRows; i++)
            {
                string name = System.Text.Encoding.UTF8.GetString(rowsOfQueryRows.Value.ARow[i].LpProps[1].Value.LpszA);

                // Server will ignore cases when comparing string according to section 2.2.6 in Open Specification MS-OXNSPI.
                if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(administratorName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    PermanentEntryID administratorEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.Bin.Lpb);
                    administratorESSDN = administratorEntryID.DistinguishedName;
                }
                else if (name.ToLower(System.Globalization.CultureInfo.CurrentCulture).Equals(mailUserName.ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                {
                    PermanentEntryID mailUserEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(rowsOfQueryRows.Value.ARow[i].LpProps[0].Value.Bin.Lpb);
                    mailUserNameESSDN = mailUserEntryID.DistinguishedName;
                }

                if (!string.IsNullOrEmpty(administratorESSDN) && !string.IsNullOrEmpty(mailUserNameESSDN))
                {
                    break;
                }
            }
            #endregion

            #region Call NspiDNToMId method to map the DNs to a set of Minimal Entry ID.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r();
            names.CValues = 2;
            names.LppszA = new string[2];
            names.LppszA[0] = administratorESSDN;
            names.LppszA[1] = mailUserNameESSDN;
            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            #endregion

            #region Call NspiGetMatches using the MID gotten in previous step to get the display name of the object.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;
            Restriction_r? filter = null;
            PropertyTagArray_r prop = new PropertyTagArray_r();
            prop.AulPropTag = new uint[]
            {
                (uint)AulProp.PidTagDisplayName
            };
            prop.CValues = (uint)prop.AulPropTag.Length;
            PropertyTagArray_r? propTagsOfGetMatches = prop;
            PropertyName_r? propNameOfGetMatches = null;

            // Output parameters.
            PropertyTagArray_r? outMIds;
            PropertyRowSet_r? rowsOfGetMatches;

            STAT statGetMatches = new STAT();
            statGetMatches.InitiateStat();
            uint administratorMID = mids.Value.AulPropTag[0];
            statGetMatches.CurrentRec = administratorMID;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref statGetMatches, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success.");
            Site.Assert.AreEqual<int>(1, rowsOfGetMatches.Value.ARow.Length, "Only one address book object should be matched.");

            string displayName = System.Text.Encoding.UTF8.GetString(rowsOfGetMatches.Value.ARow[0].LpProps[0].Value.LpszA);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R347");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R347
            // The Minimal Entry ID of the object the server is to read values from is specified in the CurrentRec field of the input parameter pStat.
            // If the two display name matched, MS-OXNSPI_R347 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                administratorName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                displayName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                347,
                @"[In MinimalEntryID] A Minimal Entry ID is a single DWORD value that identifies a specific object in the address book.");

            // As the MinimalEntryID is a Minimal Identifier and MS-OXNSPI_R347 has been verified, MS-OXNSPI_R609 can be verified directly.
            Site.CaptureRequirement(
                609,
                @"[In Object Identity] Minimal Identifier: Specifies a specific object in a single NSPI session.");
            #endregion

            #endregion

            #region Call NspiUnbind method to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the Permanent Entry ID.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC38_PermanentEntryIDVerification()
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

            // Extract the dictionary of display name and PermanentEntryID from the returned rows in the first session.
            Dictionary<string, PermanentEntryID?> dict1 = AdapterHelper.ExtractPermanentEntryIDAndDisplayname(rows);

            #endregion

            #region Call NspiUnbind with the Reserved field set to a different value.
            uint reserved = 0;
            uint returnValue2 = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>((uint)1, returnValue2, "NspiUnbind method should return 1 (Success)");
            #endregion

            #region Call NspiBind to initiate a session between the client and the server.
            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind method should return Success.");
            #endregion

            #region Call NspiGetSpecialTable method with dwFlags set to "NspiAddressCreationTemplates".

            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable method should return Success.");

            // Extract the dictionary of display name and PermanentEntryID from the returned rows in the second session.
            Dictionary<string, PermanentEntryID?> dict2 = AdapterHelper.ExtractPermanentEntryIDAndDisplayname(rows);

            #region Capture
            bool isR2001Verified = false;
            if (dict1.Count == dict2.Count)
            {
                foreach (string displayName in dict1.Keys)
                {
                    PermanentEntryID? permanentEntryID1 = dict1[displayName];
                    PermanentEntryID? permanentEntryID2 = dict2[displayName];
                    if (AdapterHelper.AreTwoPermanentEntryIDEqual(permanentEntryID1, permanentEntryID2))
                    {
                        isR2001Verified = true;
                    }
                    else
                    {
                        isR2001Verified = false;
                        break;
                    }
                }
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R2001
            Site.CaptureRequirementIfIsTrue(
                isR2001Verified,
                2001,
                "[In Object Identity] The Permanent Identifiers of a specific object are same in two NSPI sessions.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetMatches operation with different Delta and ContainerID fields.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S02_TC39_GetMatchesWithDifferentDeltaAndContainerID()
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

            #region Call NspiQueryRows to get the DN of the specified user.
            uint flagsOfQueryRows = (uint)RetrievePropertyFlag.fSkipObjects;
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
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiQueryRows should return Success!");

            string userName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
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

            #region Call NspiDNToMId to get the MIDs of the specified user.
            uint reserved = 0;
            StringsArray_r names = new StringsArray_r
            {
                LppszA = new string[]
                {
                    userESSDN
                }
            };
            names.CValues = (uint)names.LppszA.Length;
            PropertyTagArray_r? mids;
            this.Result = this.ProtocolAdatper.NspiDNToMId(reserved, names, out mids);
            #endregion

            #region Call NspiGetMatches twice with different Delta and ContainerID fields in input parameters.
            uint reserved1 = 0;
            uint reserver2 = 0;
            PropertyTagArray_r? proReserved = null;
            uint requested = Constants.GetMatchesRequestedRowNumber;
            Restriction_r? filter = null;
            PropertyTagArray_r? outMIds1;
            stat.CurrentRec = mids.Value.AulPropTag[0];
            stat.Delta = 1;
            stat.ContainerID = 1;

            PropertyTagArray_r propTag = new PropertyTagArray_r
            {
                CValues = 3,
                AulPropTag = new uint[3]
                {
                    (uint)AulProp.PidTagEntryId,
                    (uint)AulProp.PidTagDisplayType,
                    (uint)AulProp.PidTagObjectType
                }
            };

            PropertyTagArray_r? propTagsOfGetMatches = propTag;
            PropertyName_r? propNameOfGetMatches = null;

            // Output parameters.
            PropertyRowSet_r? rowsOfGetMatches;

            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds1, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            Site.Assert.IsNotNull(outMIds1.Value, "The Minimal Entry IDs returned successfully. The MId number is {0}.", outMIds1 == null ? 0 : outMIds1.Value.CValues);

            stat.Delta = 2;
            stat.ContainerID = (uint)AulProp.PidTagDisplayName;
            PropertyTagArray_r? outMIds2;
            this.Result = this.ProtocolAdatper.NspiGetMatches(reserved1, ref stat, proReserved, reserver2, filter, propNameOfGetMatches, requested, out outMIds2, propTagsOfGetMatches, out rowsOfGetMatches);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetMatches should return Success!");
            Site.Assert.IsNotNull(outMIds2.Value, "The Minimal Entry IDs returned successfully. The MId number is {0}.", outMIds2 == null ? 0 : outMIds2.Value.CValues);

            Site.Assert.AreEqual<uint>(outMIds1.Value.CValues, outMIds2.Value.CValues, "The Minimal Entry IDs count should be same.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1131");

            for (uint i = 0; i < outMIds1.Value.CValues; i++)
            {
                Site.Assert.AreEqual<uint>(outMIds1.Value.AulPropTag[i], outMIds2.Value.AulPropTag[i], "The Minimal Entry ID should be same.");
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1131
            // The Delta and ContainerID fields are different but the server returned the same Minimal Entry IDs, MS-OXNSPI_R1131 can be verified directly.
            Site.CaptureRequirement(
                1131,
                @"[In NspiGetMatches] [Server Processing Rules: Upon receiving message NspiGetMatches, the server MUST process the data from the message subject to the following constraints:] [Constraint 9] The server MUST ignore any values of the Delta and ContainerID fields while locating the object.");
            #endregion

            #region Call NspiUnbind to destroy the session between the client and the server.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        #region Private method
        /// <summary>
        /// Verify whether uses the list specified by pPropTags.
        /// </summary>
        /// <param name="propTags">The PropertyRow_r must following this.</param>
        /// <param name="propRow">It must follows the PropertyTagArray_r.</param>
        private void VerifyServerUsesListSpecifiedBypPropTagsInNspiGetProps(PropertyTagArray_r? propTags, PropertyRow_r? propRow)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R886");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R886
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsRowSubjectToPropTags(propTags, propRow),
                886,
                @"[In NspiGetProps] [Constraint 5] [If the input parameter pPropTags is not NULL] The server MUST use this list [input parameter pPropTags].");
        }

        /// <summary>
        /// Verify rows returned from NspiQueryRows is not null.
        /// </summary>
        /// <param name="rowsOfQueryRows">Contains the address book container rows that the server returns in response to the request.</param>
        private void VerifyRowsReturnedFromNspiQueryRowsIsNotNull(PropertyRowSet_r? rowsOfQueryRows)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R973");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R973
            // The rowsOfQueryRows is returned from the server which contains a RowSet, if it is not null, it illustrates that the server 
            // must have constructed it.
            Site.CaptureRequirementIfIsNotNull(
                rowsOfQueryRows,
                973,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] The server constructs a RowSet.");
        }

        /// <summary>
        /// Verify server used the list specified by the input parameter pPropTags.
        /// </summary>
        /// <param name="rowsOfQueryRows">Contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="propTags">The input list specified by pPropTags.</param>
        private void VerifyServerUsedTheListInNspiQueryRows(PropertyRowSet_r? rowsOfQueryRows, PropertyTagArray_r? propTags)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R951");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R951
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsRowSetSubjectToPropTags(propTags, rowsOfQueryRows),
                951,
                @"[In NspiQueryRows] [Server Processing Rules: Upon receiving message NspiQueryRows, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] [If the input parameter pPropTags is not NULL] The server MUST use this list [the list specified by pPropTags].");
        }

        #endregion
    }
}