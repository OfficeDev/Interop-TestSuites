namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter capture code for MS_OXCTABLAdapter
    /// </summary>
    public partial class MS_OXCTABLAdapter
    {
        #region MAPIHTTP transport

        /// <summary>
        /// Verify MAPIHTTP transport.
        /// </summary>
        private void VerifyMAPITransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && Common.IsRequirementEnabled(1340, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R1340");

                // Verify requirement MS-OXCTABL_R1340
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                Site.CaptureRequirement(
                        1340,
                        @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
            }
        }
        #endregion

        /// <summary>
        /// Verify RopSetColumns Response
        /// </summary>
        /// <param name="setColumnsResponse">RopSetColumnsResponse structure data that needs verification</param>
        private void VerifyRopSetColumnsResponse(RopSetColumnsResponse setColumnsResponse)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R54");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R54
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                setColumnsResponse.TableStatus.GetType(),
                54,
                @"[In RopSetColumns ROP Response Buffer] TableStatus (1 byte): An enumeration that indicates the status of asynchronous operations being performed on the table.<4>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R55: the current table status is {0}", setColumnsResponse.TableStatus);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R55
            bool isVerifyR55 = (setColumnsResponse.TableStatus == 0x00) ||
                               (setColumnsResponse.TableStatus == 0x09) ||
                               (setColumnsResponse.TableStatus == 0x0A) ||
                               (setColumnsResponse.TableStatus == 0x0B) ||
                               (setColumnsResponse.TableStatus == 0x0D) ||
                               (setColumnsResponse.TableStatus == 0x0E) ||
                               (setColumnsResponse.TableStatus == 0x0F);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR55,
                55,
                @"[In RopSetColumns ROP Response Buffer] It [TableStatus] MUST have one of the table status values [0x00,0x09,0x0A,0x0B,0x0D,0x0E,0x0F] that are specified in section 2.2.2.1.3.");

            // Table Status is TBLSTAT_COMPLETE with value 0x00 means no operations are in progress.
            if ((TableRopReturnValues)setColumnsResponse.ReturnValue == TableRopReturnValues.success && !this.globalIsSetColumnsAsynchronous)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R33");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R33
                Site.CaptureRequirementIfAreEqual(
                    0x00,
                    setColumnsResponse.TableStatus,
                    33,
                    @"[In TableStatus] When the Table Status is TBLSTAT_COMPLETE with value 0x00 means no operations are in progress.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R43");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R43
                Site.CaptureRequirementIfAreEqual(
                    0x00,
                    setColumnsResponse.TableStatus,
                    43,
                    @"[In Asynchronous Flags] When TBL_ASYNC value is 0x00,the server will perform the operation synchronously.");
            }

            if (this.globalIsSetColumnsAsynchronous)
            {
                if (Common.IsRequirementEnabled(817, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R817");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R817
                    Site.CaptureRequirementIfAreEqual(
                        0x0B,
                        setColumnsResponse.TableStatus,
                        817,
                        @"[In Appendix A: Product Behavior] When the Table Status returned by the implementation is TBLSTAT_SETTING_COLS with value 0x0B, it means that a RopSetColumns ROP is in progress. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(42, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R42");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R42
                    Site.CaptureRequirementIfAreEqual(
                        0x0B,
                        setColumnsResponse.TableStatus,
                        42,
                        @"[In Appendix A: Product Behavior] When TBL_ASYNC value is 0x01,the implementation performs the ROP asynchronously. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                // The requirement 812 can not be verified in the environments that cannot support asynchronous.
                if (Common.IsRequirementEnabled(812, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R812");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R812
                    Site.CaptureRequirementIfAreEqual(
                        0x00,
                        setColumnsResponse.TableStatus,
                        812,
                        @"[In Appendix A: Product Behavior] If the implementation does not honor requests to perform operations asynchronously, it will return Table Status ""TBLSTAT_COMPLETE"" with value 0x00 when asynchronous flags is set to 0x01. (Microsoft Exchange Server 2010 and above follow this behavior.)");
                }

                // If the client requests that the server perform a RopSetColumns ROP request asynchronously, the server does not perform the operation synchronously and return "TBLSTAT_COMPLETE" in the TableStatus field of the ROP response buffer
                if (Common.IsRequirementEnabled(795, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R795");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R795
                    Site.CaptureRequirementIfAreNotEqual(
                        0x00,
                        setColumnsResponse.TableStatus,
                        795,
                        @"[In Appendix A: Product Behavior] If the client requests that the implementation perform a RopSetColumns ([MS-OXCROPS] section 2.2.5.1) ROP request asynchronously, it does not perform the operation synchronously and not return ""TBLSTAT_COMPLETE"" in the TableStatus field of the ROP response buffer. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                // If the client requests that the server perform a RopSetColumns ([MS-OXCROPS] section 2.2.5.1) ROP request asynchronously, the server does perform the operation synchronously and return "TBLSTAT_COMPLETE" in the TableStatus field of the ROP response buffer
                if (Common.IsRequirementEnabled(796, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R796");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R796
                    Site.CaptureRequirementIfAreEqual(
                        0x00,
                        setColumnsResponse.TableStatus,
                        796,
                        @"[In Appendix A: Product Behavior] If the client requests that the implementation perform a RopSetColumns ([MS-OXCROPS] section 2.2.5.1) ROP request asynchronously, it does perform the operation synchronously and return ""TBLSTAT_COMPLETE"" in the TableStatus field of the ROP response buffer. ( <20> Section 3.2.5.1: Exchange 2010, Exchange 2013, and Exchange 2016 do not support asynchronous operations on tables and ignore the TABL_ASYNC flags, as described in section 2.2.2.1.4.)");
                }

                // If the TBL_ASYNC bit of the SetColumnsFlags field is set, the server can execute the ROP as a table-asynchronous ROP
                // In asynchronous ROP, the TableStatus field in response is not TBLSTAT_COMPLETE.
                if (Common.IsRequirementEnabled(793, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R793");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R793
                    Site.CaptureRequirementIfAreNotEqual(
                        0x00,
                        setColumnsResponse.TableStatus,
                        793,
                        @"[In Appendix A: Product Behavior] If the TBL_ASYNC bit of the SetColumnsFlags field is set, the implementation can execute the ROP as a table-asynchronous ROP, as specified in section 3.2.5.1. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(770, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R770");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R770
                    Site.CaptureRequirementIfAreEqual(
                        0x00,
                        setColumnsResponse.TableStatus,
                        770,
                        @"[In Appendix A: Product Behavior] Implementation does return Table Status ""TBLSTAT_COMPLETE"" with value 0x00 when asynchronous flags is set to 0x01 in a RopSetColumns ROP on tables. (<3> Section 2.2.2.2.1: Exchange 2010, Exchange 2013, and Exchange 2016 will return Table Status ""TBLSTAT_COMPLETE"" with value 0x00 when asynchronous flags is set to 0x01 in a RopSetColumns ROP on tables.)");
                }

                // Here it can partially verify MS-OXCTABL requirement: MS-OXCTABL_418
                // if the server executes the ROP asynchronously, the server return "TBLSTAT_SETTING_COLS" in the TableStatus field of the ROP response buffer and do the work asynchronously.
                if (Common.IsRequirementEnabled(418, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R418");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R418
                    Site.CaptureRequirementIfAreEqual(
                        0x0B,
                        setColumnsResponse.TableStatus,
                        418,
                        @"[In Appendix A: Product Behavior] However, if the implementation executes the ROP asynchronously, it return ""TBLSTAT_SORTING"", ""TBLSTAT_SETTING_COLS"", or ""TBLSTAT_RESTRICTING"" (depending on the ROP performed) in the TableStatus field of the ROP response buffer and do the work asynchronously. (Microsoft Exchange Server 2007 follows this behavior.)");
                }
            }
        }

        /// <summary>
        /// Verify RopSortTable Response
        /// </summary>
        /// <param name="sortTableResponse">RopSortTableResponse structure data that needs verification</param>
        private void VerifyRopSortTableResponse(RopSortTableResponse sortTableResponse)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R83");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R83
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                sortTableResponse.TableStatus.GetType(),
                83,
                @"[In RopSortTable ROP Response Buffer] TableStatus (1 byte): An enumeration that indicates the status of asynchronous operations being performed on the table.<6>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R84: the current table status is {0}", sortTableResponse.TableStatus);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R84
            bool isVerifyR84 = (sortTableResponse.TableStatus == 0x00) ||
                               (sortTableResponse.TableStatus == 0x09) ||
                               (sortTableResponse.TableStatus == 0x0A) ||
                               (sortTableResponse.TableStatus == 0x0B) ||
                               (sortTableResponse.TableStatus == 0x0D) ||
                               (sortTableResponse.TableStatus == 0x0E) ||
                               (sortTableResponse.TableStatus == 0x0F);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR84,
                84,
                @"[In RopSortTable ROP Response Buffer] This field MUST have one of the table status values[0x00,0x09,0x0A,0x0B,0x0D,0x0E,0x0F] that are specified in section 2.2.2.1.3.");

            // Table Status is TBLSTAT_COMPLETE with value 0x00 means no operations are in progress.
            if ((TableRopReturnValues)sortTableResponse.ReturnValue == TableRopReturnValues.success && !this.globalIsSortTableAsynchronous)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R33");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R33
                Site.CaptureRequirementIfAreEqual(
                    0x00,
                    sortTableResponse.TableStatus,
                    33,
                    @"[In TableStatus] When the Table Status is TBLSTAT_COMPLETE with value 0x00 means no operations are in progress.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R43");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R43
                Site.CaptureRequirementIfAreEqual(
                    0x00,
                    sortTableResponse.TableStatus,
                    43,
                    @"[In Asynchronous Flags] When TBL_ASYNC value is 0x00,the server will perform the operation synchronously.");
            }

            if (this.globalIsSortTableAsynchronous)
            {
                if (Common.IsRequirementEnabled(815, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R815");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R815
                    Site.CaptureRequirementIfAreEqual(
                        0x09,
                        sortTableResponse.TableStatus,
                        815,
                        @"[In Appendix A: Product Behavior] When the Table Status returned by the implementation is TBLSTAT_SORTING with value 0x09 means a RopSortTable ROP is in progress. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(42, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R42");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R42
                    Site.CaptureRequirementIfAreEqual(
                        0x09,
                        sortTableResponse.TableStatus,
                        42,
                        @"[In Appendix A: Product Behavior] When TBL_ASYNC value is 0x01,the implementation performs the ROP asynchronously. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(799, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R799");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R799
                    Site.CaptureRequirementIfAreNotEqual(
                        0x00,
                        sortTableResponse.TableStatus,
                        799,
                        @"[In Appendix A: Product Behavior] If the client requests that the implementation perform a RopSortTable ([MS-OXCROPS] section 2.2.5.2) ROP request asynchronously, it does not perform the operation synchronously and not return ""TBLSTAT_COMPLETE"" in the TableStatus field of the ROP response buffer. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(800, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R800");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R800
                    Site.CaptureRequirementIfAreEqual(
                        0x00,
                        sortTableResponse.TableStatus,
                        800,
                        @"[In Appendix A: Product Behavior] If the client requests that the implementation perform a RopSortTable ([MS-OXCROPS] section 2.2.5.2) ROP request asynchronously, it does perform the operation synchronously and return ""TBLSTAT_COMPLETE"" in the TableStatus field of the ROP response buffer. ( <20> Section 3.2.5.1: Exchange 2010, Exchange 2013, and Exchange 2016 do not support asynchronous operations on tables and ignore the TABL_ASYNC flags, as described in section 2.2.2.1.4.)");
                }

                // If the TBL_ASYNC bit of the SortTableFlags field is set, the server can execute the ROP as a table-asynchronous ROP
                // In asynchronous ROP, the TableStatus field in response is not TBLSTAT_COMPLETE.
                if (Common.IsRequirementEnabled(794, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R794");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R794
                    Site.CaptureRequirementIfAreNotEqual(
                        0x00,
                        sortTableResponse.TableStatus,
                        794,
                        @"[In Appendix A: Product Behavior] If the TBL_ASYNC bit of the SortTableFlags field is set, the implementation can execute the ROP as a table-asynchronous ROP,as specified in section 3.2.5.1. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                // In the environments that cannot support asynchronous, it will ignore the asynchronous flags, and return Table Status "TBLSTAT_COMPLETE" with value 0x00 when asynchronous flags is set to 0x01 in a RopSortTable ROP.
                if (Common.IsRequirementEnabled(772, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R772");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R772
                    Site.CaptureRequirementIfAreEqual(
                        0x00,
                        sortTableResponse.TableStatus,
                        772,
                        @"[In Appendix A: Product Behavior] Implementation does return Table Status ""TBLSTAT_COMPLETE"" with value 0x00 when asynchronous flags is set to 0x01 in a RopSortTable ROP on tables. (<5> Section 2.2.2.3.1: Exchange 2010, Exchange 2013, and Exchange 2016 will return Table Status ""TBLSTAT_COMPLETE"" with value 0x00 when asynchronous flags is set to 0x01 in a RopSortTable ROP on tables.)");
                }

                // Here it can partially verify MS-OXCTABL requirement: MS-OXCTABL_418
                // if the server executes the ROP asynchronously, the server return "TBLSTAT_SORTING" in the TableStatus field of the ROP response buffer and do the work asynchronously.
                if (Common.IsRequirementEnabled(418, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R418");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R418
                    Site.CaptureRequirementIfAreEqual(
                        0x09,
                        sortTableResponse.TableStatus,
                        418,
                        @"[In Appendix A: Product Behavior] However, if the implementation executes the ROP asynchronously, it return ""TBLSTAT_SORTING"", ""TBLSTAT_SETTING_COLS"", or ""TBLSTAT_RESTRICTING"" (depending on the ROP performed) in the TableStatus field of the ROP response buffer and do the work asynchronously. (Microsoft Exchange Server 2007 follows this behavior.)");
                }
            }
        }

        /// <summary>
        /// Verify RopRestrict Response
        /// </summary>
        /// <param name="restrictResponse">RopRestrictResponse structure data that needs verification</param>
        private void VerifyRopRestrictResponse(RopRestrictResponse restrictResponse)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R93");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R93
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                restrictResponse.TableStatus.GetType(),
                93,
                @"[In RopRestrict ROP Response Buffer] TableStatus (1 byte): An enumeration that indicates the status of asynchronous operations being performed on the table.<8>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R94: the current table status is {0}", restrictResponse.TableStatus);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R94
            bool isVerifyR94 = (restrictResponse.TableStatus == 0x00) ||
                               (restrictResponse.TableStatus == 0x09) ||
                               (restrictResponse.TableStatus == 0x0A) ||
                               (restrictResponse.TableStatus == 0x0B) ||
                               (restrictResponse.TableStatus == 0x0D) ||
                               (restrictResponse.TableStatus == 0x0E) ||
                               (restrictResponse.TableStatus == 0x0F);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR94,
                94,
                @"[In TableStatus] This field MUST have one of the table status values[0x00,0x09,0x0A,0x0B,0x0D,0x0E,0x0F] that are specified in section 2.2.2.1.3.");

            if ((TableRopReturnValues)restrictResponse.ReturnValue == TableRopReturnValues.success && !this.globalIsRestrictAsynchronous)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R33");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R33
                Site.CaptureRequirementIfAreEqual(
                    0x00,
                    restrictResponse.TableStatus,
                    33,
                    @"[In TableStatus] When the Table Status is TBLSTAT_COMPLETE with value 0x00 means no operations are in progress.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R43");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R43
                Site.CaptureRequirementIfAreEqual(
                    0x00,
                    restrictResponse.TableStatus,
                    43,
                    @"[In Asynchronous Flags] When TBL_ASYNC value is 0x00,the server will perform the operation synchronously.");
            }

            if (this.globalIsRestrictAsynchronous)
            {
                if (Common.IsRequirementEnabled(819, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R819");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R819
                    Site.CaptureRequirementIfAreEqual(
                        0x0E,
                        restrictResponse.TableStatus,
                        819,
                        @"[In Appendix A: Product Behavior] When the Table Status returned by the implementation is TBLSTAT_RESTRICTING with value 0x0E means a RopRestrict ROP is in progress. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(42, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R42");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R42
                    Site.CaptureRequirementIfAreEqual(
                        0x0E,
                        restrictResponse.TableStatus,
                        42,
                        @"[In Appendix A: Product Behavior] When TBL_ASYNC value is 0x01,the implementation performs the ROP asynchronously. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                // If the client requests that the server perform a RopRestrict ROP request asynchronously, the server does not perform the operation synchronously and return "TBLSTAT_COMPLETE" in the TableStatus field of the ROP response buffer.
                if (Common.IsRequirementEnabled(801, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R801");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R801
                    Site.CaptureRequirementIfAreNotEqual(
                        0x00,
                        restrictResponse.TableStatus,
                        801,
                        @"[In Appendix A: Product Behavior] If the client requests that the implementation perform a RopRestrict ([MS-OXCROPS] section 2.2.5.3) ROP request asynchronously, it does not perform the operation synchronously and not return ""TBLSTAT_COMPLETE"" in the TableStatus field of the ROP response buffer. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(811, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R811");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R811
                    Site.CaptureRequirementIfAreNotEqual(
                        0x00,
                        restrictResponse.TableStatus,
                        811,
                        @"[In Appendix A: Product Behavior] If the TBL_ASYNC bit of the RestrictFlags field is set, the implementation does execute the ROP as a table-asynchronous ROP, as specified in section 3.2.5.1. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(802, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R802");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R802
                    Site.CaptureRequirementIfAreEqual(
                        0x00,
                        restrictResponse.TableStatus,
                        802,
                        @"[In Appendix A: Product Behavior] If the client requests that the implementation perform a RopRestrict ([MS-OXCROPS] section 2.2.5.3) ROP request asynchronously, it does perform the operation synchronously and return ""TBLSTAT_COMPLETE"" in the TableStatus field of the ROP response buffer. ( <20> Section 3.2.5.1: Exchange 2010, Exchange 2013, and Exchange 2016 do not support asynchronous operations on tables and ignore the TABL_ASYNC flags, as described in section 2.2.2.1.4.)");
                }

                // In the environments that cannot support asynchronous, it will ignore the asynchronous flags, and return Table Status "TBLSTAT_COMPLETE" with value 0x00 when asynchronous flags is set to 0x01 in a RopRestrict ROP.
                if (Common.IsRequirementEnabled(774, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R774");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R774
                    Site.CaptureRequirementIfAreEqual(
                        0x00,
                        restrictResponse.TableStatus,
                        774,
                        @"[In Appendix A: Product Behavior] Implementation does return Table Status ""TBLSTAT_COMPLETE"" with value 0x00 when asynchronous flags is set to 0x01 in a RopRestrict ROP on tables. (<7> Section 2.2.2.4.1: Exchange 2010, Exchange 2013, and Exchange 2016 will return Table Status ""TBLSTAT_COMPLETE"" with value 0x00 when asynchronous flags is set to 0x01 in a RopRestrict ROP on tables.)");
                }

                // Here it can partially verify MS-OXCTABL requirement: MS-OXCTABL_418
                // if the server executes the ROP asynchronously, the server return "TBLSTAT_RESTRICTING" in the TableStatus field of the ROP response buffer and do the work asynchronously.
                if (Common.IsRequirementEnabled(418, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R418");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R418
                    Site.CaptureRequirementIfAreEqual(
                        0x0E,
                        restrictResponse.TableStatus,
                        418,
                        @"[In Appendix A: Product Behavior] However, if the implementation executes the ROP asynchronously, it return ""TBLSTAT_SORTING"", ""TBLSTAT_SETTING_COLS"", or ""TBLSTAT_RESTRICTING"" (depending on the ROP performed) in the TableStatus field of the ROP response buffer and do the work asynchronously. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(810, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R810");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R810
                    // TableStatus in response with value 0x00 means server executes synchronous ROP  
                    Site.CaptureRequirementIfAreEqual<byte>(
                        0x00,
                        restrictResponse.TableStatus,
                        810,
                        @"[In Appendix A: Product Behavior] If the TBL_ASYNC bit of the RestrictFlags field is set, the implementation does not execute the ROP as a table-asynchronous ROP, as specified in section 3.2.5.1. (<24> Section 3.2.5.4: Exchange 2010, Exchange 2013, and Exchange 2016 do not support asynchronous operations on tables and ignore the TABL_ASYNC flags, as described in section 2.2.2.1.4.)");
                }
            }
        }

        /// <summary>
        /// Verify RopQueryRows Response
        /// </summary>
        /// <param name="queryRowsResponse">RopQueryRowsResponse structure data that needs verification</param>
        /// <param name="rowCountRequest">The RowCount that is specified in the request</param>
        private void VerifyRopQueryRowsResponse(RopQueryRowsResponse queryRowsResponse, ushort rowCountRequest)
        {
            this.VerifyRPCLayerRequirement();

            if (queryRowsResponse.RowCount > 0)
            {
                // The value of the following two properties are only valid for content table.
                if (this.tableType == TableType.CONTENT_TABLE)
                {
                    if (queryRowsResponse.RowData.PropertyRows != null)
                    {
                        for (int i = 0; i < queryRowsResponse.RowData.PropertyRows.Count; i++)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R617");

                            // Verify MS-OXCTABL requirement: MS-OXCTABL_R617
                            Site.CaptureRequirementIfAreEqual<int>(
                                8,
                                queryRowsResponse.RowData.PropertyRows[i].PropertyValues[0].Value.Length,
                                617,
                                @"[In PidTagInstID] Data type: PtypInteger64 ([MS-OXCDATA] section 2.11.1).");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R618");

                            // Verify MS-OXCTABL requirement: MS-OXCTABL_R618
                            Site.CaptureRequirementIfAreEqual<int>(
                                4,
                                queryRowsResponse.RowData.PropertyRows[i].PropertyValues[1].Value.Length,
                                618,
                                @"[In PidTagInstanceNum] Data type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R619");

                            // Verify MS-OXCTABL requirement: MS-OXCTABL_R619
                            Site.CaptureRequirementIfAreEqual<int>(
                                4,
                                queryRowsResponse.RowData.PropertyRows[i].PropertyValues[3].Value.Length,
                                619,
                                @"[In PidTagRowType] Data type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R620");

                            // Verify MS-OXCTABL requirement: MS-OXCTABL_R620
                            Site.CaptureRequirementIfAreEqual<int>(
                                4,
                                queryRowsResponse.RowData.PropertyRows[i].PropertyValues[4].Value.Length,
                                620,
                                @"[In PidTagDepth] Data type: PtypInteger32 property ([MS-OXCDATA] section 2.11.1).");
                        }
                    }

                    if (!this.areMultipleSortOrders && this.isExpanded == true)
                    {
                        int i = 0;
                        uint tempRowType = 0;
                        bool isLeafRowExist = false;
                        for (; i < queryRowsResponse.RowData.PropertyRows.Count; i++)
                        {
                            tempRowType = BitConverter.ToUInt32(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[3].Value, 0);
                            if (tempRowType == 0x01)
                            {
                                isLeafRowExist = true;
                                break;
                            }
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R15{0}", isLeafRowExist ? string.Empty : string.Format("There is no leaf row in the returned property rows!"));

                        // Verify MS-OXCTABL requirement: MS-OXCTABL_R15
                        Site.CaptureRequirementIfIsTrue(
                            isLeafRowExist,
                            15,
                            @"[In PidTagRowType] When the PidTagRowType is TBL_LEAF_ROW with value 0x00000001 means the row is a row of data.");
                    }

                    if (this.isExpanded == true)
                    {
                        int i = 0;
                        uint tempRowType = 0;
                        bool isExpandedHeaderExist = false;
                        for (; i < queryRowsResponse.RowData.PropertyRows.Count; i++)
                        {
                            tempRowType = BitConverter.ToUInt32(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[3].Value, 0);
                            if (tempRowType == 0x03)
                            {
                                isExpandedHeaderExist = true;
                                break;
                            }
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R17{0}", isExpandedHeaderExist ? string.Empty : string.Format("There is no expanded header row in the returned property rows!"));

                        // Verify MS-OXCTABL requirement: MS-OXCTABL_R17
                        Site.CaptureRequirementIfIsTrue(
                            isExpandedHeaderExist,
                            17,
                            @"[In PidTagRowType] When the PidTagRowType is TBL_EXPANDED_CATEGORY with value 0x00000003 means the row is a header row that is expanded.");
                    }
                    else if ((!this.areMultipleSortOrders && this.areAllSortOrdersUsedAsCategory) || (this.areMultipleSortOrders && !this.areAllSortOrdersUsedAsCategory))
                    {
                        int i = 0;
                        uint tempRowType = 0;
                        bool isCollapsedHeaderExist = false;
                        for (; i < queryRowsResponse.RowData.PropertyRows.Count; i++)
                        {
                            tempRowType = BitConverter.ToUInt32(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[3].Value, 0);
                            if (tempRowType == 0x04)
                            {
                                isCollapsedHeaderExist = true;
                                break;
                            }
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R18{0}", isCollapsedHeaderExist ? string.Empty : string.Format("There is no collapsed header row in the returned property rows!"));

                        // Verify MS-OXCTABL requirement: MS-OXCTABL_R18
                        Site.CaptureRequirementIfIsTrue(
                            isCollapsedHeaderExist,
                            18,
                            @"[In PidTagRowType] When the PidTagRowType is TBL_COLLAPSED_CATEGORY with value 0x00000004 means the row is a header row that is collapsed.");
                    }

                    // Since MS-OXCTABL_R15, MS-OXCTABL_R17 and MS-OXCTABL_R18 are verified, this requirement can be captured directly.    
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R13");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R13
                    Site.CaptureRequirement(
                        13,
                        @"[In PidTagRowType] The PidTagRowType property ([MS-OXPROPS] section 2.931) identifies the type of the row.");

                    if (this.areMultipleSortOrders && this.areAllSortOrdersUsedAsCategory && this.areAllCategoryExpanded)
                    {
                        int i = 0;
                        uint tempDepth = 0;
                        bool isCorrectDepth = false;

                        // There are no more than 2 categories in the response in this test suite.
                        for (; i < 3; i++)
                        {
                            tempDepth = BitConverter.ToUInt32(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[4].Value, 0);
                            if (tempDepth == i)
                            {
                                isCorrectDepth = true;
                            }
                            else
                            {
                                break;
                            }
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R19{0}", isCorrectDepth ? string.Empty : string.Format(": The no.{0} row depth is not correct and the value is {1}!", i, tempDepth));

                        // Verify MS-OXCTABL requirement: MS-OXCTABL_R19
                        Site.CaptureRequirementIfIsTrue(
                            isCorrectDepth,
                            19,
                            @"[In PidTagDepth] The PidTagDepth property ([MS-OXPROPS] section 2.664) specifies the number of nested categories in which a given row is contained.");
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R110: the value of the Origin field in the response for RopQueryRows is set to {0}", queryRowsResponse.Origin);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R110
            bool isVerifyR110 = (queryRowsResponse.Origin == 0x00) ||
                                (queryRowsResponse.Origin == 0x01) ||
                                (queryRowsResponse.Origin == 0x02);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR110,
                110,
                @"[In RopQueryRows ROP Response Buffer] This field [Origin] MUST be set to one of the predefined bookmark values[0x00,0x01,0x02] specified in section 2.2.2.1.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R109");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R109
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                queryRowsResponse.Origin.GetType(),
                109,
                @"[In RopQueryRows ROP Response Buffer] Origin (1 byte): An enumeration that identifies the cursor position.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R111 (This is a WORD field)");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R111
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                queryRowsResponse.RowCount.GetType(),
                111,
                @"[In RopQueryRows ROP Response Buffer] RowCount (2 bytes): An unsigned integer that specifies the number of rows returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R112", rowCountRequest);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R112
            bool isVerifyR112 = (queryRowsResponse.RowCount <= rowCountRequest) &&
                                 (queryRowsResponse.RowCount >= 0x0000);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR112,
                112,
                @"[In RopQueryRows ROP Response Buffer] Its [RowCount's] value MUST be less than or equal to the RowCount field value that is specified in the request, and it MUST be greater than or equal to 0x0000.");

            if (queryRowsResponse.RowData.PropertyRows != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R113");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R113
                Site.CaptureRequirementIfAreEqual<ushort>(
                    (ushort)queryRowsResponse.RowData.PropertyRows.Count,
                    queryRowsResponse.RowCount,
                    113,
                    @"[In RopQueryRows ROP Response Buffer] It [RowCount] MUST be equal to the number of PropertyRow objects returned in the RowData field.");

                // If the propertyRows count in the queryrowsResponse is not zero, this requirement can be covered
                if (queryRowsResponse.RowData.PropertyRows.Count > 0)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R114");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R114
                    Site.CaptureRequirement(
                        114,
                        @"[In RopQueryRows ROP Response Buffer] RowData (variable): A list of PropertyRow structures that contains the array of rows returned.");
                }

                for (int i = 0; i < queryRowsResponse.RowData.PropertyRows.Count; i++)
                {
                    int totalSize = 0;
                    for (int j = 0; j < queryRowsResponse.RowData.PropertyRows[i].PropertyValues.Count; j++)
                    {
                        if (queryRowsResponse.RowData.PropertyRows[i].PropertyValues[j].Value == null)
                        {
                            continue;
                        }

                        if (queryRowsResponse.RowData.PropertyRows[i].PropertyValues[j].Value != null)
                        {
                            totalSize += queryRowsResponse.RowData.PropertyRows[i].PropertyValues[j].Size();
                        }
                    }

                    bool isVerifyR121 = totalSize <= 510 * sizeof(byte);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R121", 510 * sizeof(byte));

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R121
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR121,
                        121,
                        @"[In RopQueryRows ROP Response Buffer] Every property value returned in a row MUST be less than or equal to 510 bytes in size.");
                }

                foreach (PropertyRow propertyRow in queryRowsResponse.RowData.PropertyRows)
                {
                    this.VerifyPropertyRowStructure(propertyRow);
                }
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R113");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R113
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0,
                    queryRowsResponse.RowCount,
                    113,
                    @"[In RopQueryRows ROP Response Buffer] It [RowCount] MUST be equal to the number of PropertyRow objects returned in the RowData field.");
            }

            for (int i = 0; i < queryRowsResponse.RowCount; i++)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R115");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R115
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(PropertyRow),
                    queryRowsResponse.RowData.PropertyRows[i].GetType(),
                    115,
                    @"[In RopQueryRows ROP Response Buffer] Each row is represented by a PropertyRow object, as specified in [MS-OXCDATA] section 2.8.1.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R463: the value of the RowCount field in the response for RopQueryRows is set to {0}", queryRowsResponse.RowCount);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R463
            bool isVerifyR463 = queryRowsResponse.RowCount <= rowCountRequest;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR463,
                463,
                @"[In Processing RopQueryRows] The number of rows sent in the ROP response MUST be less than or equal to the number of rows specified in the RowCount field.");
        }

        /// <summary>
        /// Verify RopAbort Response
        /// </summary>
        /// <param name="abortResponse">RopAbortResponse structure data that needs verification</param>
        private void VerifyRopAbortResponse(RopAbortResponse abortResponse)
        {
            this.VerifyRPCLayerRequirement();

            if (Common.IsRequirementEnabled(791, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R791");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R791
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    abortResponse.TableStatus.GetType(),
                    791,
                    @"[In Appendix A: Product Behavior] TableStatus (1 byte): An enumeration that indicates the status of asynchronous operations being performed on the table before the abort on the implementation. (Microsoft Exchange Server 2007 follows this behavior.)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R127: the value of the TableStatus field in the response for RopAbort is set to {0}", abortResponse.TableStatus);

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R127
                bool isVerifyR127 = (abortResponse.TableStatus == 0x00) ||
                                    (abortResponse.TableStatus == 0x09) ||
                                    (abortResponse.TableStatus == 0x0A) ||
                                    (abortResponse.TableStatus == 0x0B) ||
                                    (abortResponse.TableStatus == 0x0D) ||
                                    (abortResponse.TableStatus == 0x0E) ||
                                    (abortResponse.TableStatus == 0x0F);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR127,
                    127,
                    @"[In RopAbort ROP Response Buffer] Its value MUST be one of the table status values [0x00,0x09,0x0A,0x0B,0x0D,0x0E,0x0F] that are specified in section 2.2.2.1.3.");

                // If there were no asynchronous operations to abort, or the server was unable to abort the operations.
                // The error code ecUnableToAbort will be returned with value 0x80040114.
                if (abortResponse.ReturnValue == 0x80040114)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R479");

                    // Here it can partially verify MS-OXCTABL requirement: MS-OXCTABL_R479
                    Site.CaptureRequirement(
                        479,
                        @"[In Processing RopAbort] The RopAbort ROP ([MS-OXCROPS] section 2.2.5.5) MUST abort the current asynchronous table ROP that is executing on the table or send an error if there is nothing to abort or if it fails to abort.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R485");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R485
                    Site.CaptureRequirement(
                        485,
                        @"[In Processing RopAbort] The error code ecUnableToAbort will be returned with value 0x80040114,%x14.01.04.80, if there were no asynchronous operations to abort, or the server was unable to abort the operations.");
                }
            }
        }

        /// <summary>
        /// Verify GetStatus Response
        /// </summary>
        /// <param name="getStatusResponse">RopGetStatusResponse structure data that needs verification</param>
        private void VerifyGetStatusResponse(RopGetStatusResponse getStatusResponse)
        {
            this.VerifyRPCLayerRequirement();

            if (Common.IsRequirementEnabled(792, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R792");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R792
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    getStatusResponse.TableStatus.GetType(),
                    792,
                    @"[In Appendix A: Product Behavior] TableStatus (1 byte): An enumeration that indicates the status of asynchronous operations being performed on the table on the implementation. (Microsoft Exchange Server 2007 follows this behavior.)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R132: the value of the TableStatus field in the response for RopGetStatus is set to {0}", getStatusResponse.TableStatus);

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R132
                bool isVerifyR132 = (getStatusResponse.TableStatus == 0x00) ||
                                    (getStatusResponse.TableStatus == 0x09) ||
                                    (getStatusResponse.TableStatus == 0x0A) ||
                                    (getStatusResponse.TableStatus == 0x0B) ||
                                    (getStatusResponse.TableStatus == 0x0D) ||
                                    (getStatusResponse.TableStatus == 0x0E) ||
                                    (getStatusResponse.TableStatus == 0x0F);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR132,
                    132,
                    @"[In RopGetStatus ROP Response Buffer] Its [TableStatus'] value MUST be one of the table status values [0x00,0x09,0x0A,0x0B,0x0D,0x0E,0x0F] that are specified in section 2.2.2.1.3.");
            }
        }

        /// <summary>
        /// Verify ropQuery position
        /// </summary>
        /// <param name="queryPositionResponse">RopQueryPositionResponse structure data that needs verification</param>
        private void VerifyRopQueryPositionResponse(RopQueryPositionResponse queryPositionResponse)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R136", queryPositionResponse.Numerator);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R136
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                queryPositionResponse.Denominator.GetType(),
                136,
                @"[In RopQueryPosition ROP Response Buffer] Numerator (4 bytes): An unsigned integer that indicates the index (0-based) of the current row.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R137: the value of the Numerator field in the response for RopQueryPosition is set to {0}", queryPositionResponse.Numerator);
            
            // Verify MS-OXCTABL requirement: MS-OXCTABL_R137
            bool isVerifyR137 = queryPositionResponse.Numerator >= 0x00000000;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR137,
                137,
                @"[In RopQueryPosition ROP Response Buffer] Its [Numerator's] value MUST be greater than or equal to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R138", queryPositionResponse.Denominator);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R138
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                queryPositionResponse.Denominator.GetType(),
                138,
                @"[In RopQueryPosition ROP Response Buffer] Denominator (4 bytes): An unsigned integer that indicates the total number of rows in the table.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R139", queryPositionResponse.Numerator);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R139
            bool isVerifyR139 = queryPositionResponse.Denominator >= queryPositionResponse.Numerator;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR139,
                139,
                @"[In RopQueryPosition ROP Response Buffer] Its [Denominator's] value MUST be greater than or equal to the value of the Numerator field.");
        }

        /// <summary>
        /// Verify RopSeekRow Response
        /// </summary>
        /// <param name="seekRowResponse">RopSeekRowResponse structure data that needs verification</param>
        /// <param name="rowCountValueInTheRequest">The value of the RowCount field in the request</param>
        private void VerifyRopSeekRowResponse(RopSeekRowResponse seekRowResponse, long rowCountValueInTheRequest)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R153");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R153
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                seekRowResponse.HasSoughtLess.GetType(),
                153,
                @"[In RopSeekRow ROP Response Buffer] HasSoughtLess (1 byte): A Boolean that specifies whether the number of rows moved is less than the number of rows requested.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R159");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R159
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(int),
                seekRowResponse.RowsSought.GetType(),
                159,
                @"[In RopSeekRow ROP Response Buffer] RowsSought (4 bytes): A signed integer that specifies the actual number of rows moved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R156");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R156
            // That the field is not null means it is present.
            Site.CaptureRequirementIfIsNotNull(
                seekRowResponse.HasSoughtLess,
                156,
                @"[In RopSeekRow ROP Response Buffer] The HasSoughtLess field MUST be present in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R162");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R162
            // That the field is not null means it is present.
            Site.CaptureRequirementIfIsNotNull(
                seekRowResponse.RowsSought,
                162,
                @"[In RopSeekRow ROP Response Buffer] This field [RowsSought] MUST be present in the response.");

            if (rowCountValueInTheRequest > 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R161: the value of the RowsSought field in the response for RopSeekRow is set to {0}", seekRowResponse.RowsSought);

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R161
                bool isVerifyR161 = seekRowResponse.RowsSought > 0;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR161,
                    161,
                    @"[In RopSeekRow ROP Response Buffer] If the value of the RowCount field (in the request) is positive, then the value of the RowsSought field MUST also be positive.");
            }
            else if (rowCountValueInTheRequest < 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R160: the value of the RowsSought field in the response for RopSeekRow is set to {0}", seekRowResponse.RowsSought);

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R160
                bool isVerifyR160 = seekRowResponse.RowsSought <= 0;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR160,
                    160,
                    @"[In RopSeekRow ROP Response Buffer] If the value of the RowCount field (in the request) is negative, the value of the RowsSought field MUST also be negative or 0x00000000, indicating that the seek was performed backwards.");
            }
        }

        /// <summary>
        /// Verify RopSeekRowBookmark Response
        /// </summary>
        /// <param name="seekRowBookmarkResponse">RopSeekRowBookmarkResponse structure data that needs verification</param>
        /// <param name="rowCountValueInTheReques">The value of the RowCount field in the request</param>
        private void VerifyRopSeekRowBookmarkResponse(RopSeekRowBookmarkResponse seekRowBookmarkResponse, long rowCountValueInTheReques)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R181");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R181
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                seekRowBookmarkResponse.RowNoLongerVisible.GetType(),
                181,
                @"[In RopSeekRowBookmark ROP Response Buffer] RowNoLongerVisible (1 byte): A Boolean that indicates whether the row to which the bookmark pointed is no longer visible.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R186");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R186
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                seekRowBookmarkResponse.HasSoughtLess.GetType(),
                186,
                @"[In RopSeekRowBookmark ROP Response Buffer] HasSoughtLess (1 byte): A Boolean that specifies whether the number of rows moved is less than the number of rows requested.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R192");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R192
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                seekRowBookmarkResponse.RowsSought.GetType(),
                192,
                @"[In RopSeekRowBookmark ROP Response Buffer] RowsSought (4 bytes): An unsigned integer that specifies the actual number of rows moved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R189");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R189
            // Not null means be present
            Site.CaptureRequirementIfIsNotNull(
                seekRowBookmarkResponse.HasSoughtLess,
                189,
                @"[In RopSeekRowBookmark ROP Response Buffer] The HasSoughtLess field MUST be present in the response.");

            if (rowCountValueInTheReques < 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R193: the value of the RowsSought field in the response for RopSeekRowBookmark is set to {0}", seekRowBookmarkResponse.RowsSought);

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R193
                bool isVerifyR193 = (seekRowBookmarkResponse.RowsSought < 0)
                                   || (seekRowBookmarkResponse.RowsSought == 0x00000000);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR193,
                    193,
                    @"[In RopSeekRowBookmark ROP Response Buffer] If the value of the RowCount field in the request is negative, the value of the RowsSought field MUST also be negative or zero (0x00000000), indicating that the seek was performed backwards.");
            }

            if (rowCountValueInTheReques > 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R194: the value of the RowsSought field in the response for RopSeekRowBookmark is set to {0}", seekRowBookmarkResponse.RowsSought);

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R194
                bool isVerifyR194 = seekRowBookmarkResponse.RowsSought > 0;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR194,
                    194,
                    @"[In RopSeekRowBookmark ROP Response Buffer] If the value of the RowCount field in the request is positive, the value of RowsSought MUST also be positive.");
            }
        }

        /// <summary>
        /// Verify RopCreateBookmark Response
        /// </summary>
        /// <param name="createBookmarkResponse">RopCreateBookmarkResponse structure data that needs verification</param>
        private void VerifyRopCreateBookmarkResponse(RopCreateBookmarkResponse createBookmarkResponse)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R211 ");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R211
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                createBookmarkResponse.BookmarkSize.GetType(),
                211,
                @"[In RopCreateBookmark ROP Response Buffer] BookmarkSize (2 bytes): An unsigned integer that specifies the size, in bytes, of the Bookmark field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R213 ");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R213
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[]),
                createBookmarkResponse.Bookmark.GetType(),
                213,
                @"[In RopCreateBookmark ROP Response Buffer] Bookmark (variable): An array of bytes that contains the bookmark data.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R212");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R212
            // Not null means be present.
            Site.CaptureRequirementIfIsNotNull(
                createBookmarkResponse.BookmarkSize.GetType(),
                212,
                @"[In RopCreateBookmark ROP Response Buffer] This field [BookmarkSize] MUST be present.");
        }

        /// <summary>
        /// Verify RopQueryColumnsALL Response
        /// </summary>
        /// <param name="queryColumnsAllResponse">RopQueryColumnsAllResponse structure data that needs verification</param>
        private void VerifyRopQueryColumnsALLResponse(RopQueryColumnsAllResponse queryColumnsAllResponse)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R220");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R220
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                queryColumnsAllResponse.PropertyTagCount.GetType(),
                220,
                @"[In RopQueryColumnsAll ROP Response Buffer] PropertyTagCount (2 bytes): An unsigned integer that specifies the number of property tags in the PropertyTags field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R221");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R221
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyTag[]),
                queryColumnsAllResponse.PropertyTags.GetType(),
                221,
                @"[In RopQueryColumnsAll ROP Response Buffer] PropertyTags (variable): An array of property tags, each of which corresponds to an available column in the table.");

            for (int i = 0; i < queryColumnsAllResponse.PropertyTags.Length; i++)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R222");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R222
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(PropertyTag),
                    queryColumnsAllResponse.PropertyTags[i].GetType(),
                    222,
                    @"[In RopQueryColumnsAll ROP Response Buffer] Each property tag is represented by a PropertyTag structure, as specified in [MS-OXCDATA] section 2.9.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R181");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R181
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    queryColumnsAllResponse.PropertyTags[i].PropertyType.GetType(),
                    "MS-OXCDATA",
                    181,
                    @"[In PropertyTag Structure] PropertyType (2 bytes): An unsigned integer that identifies the data type of the property value, as specified by the table in section 2.11.1.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R182");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R182
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    queryColumnsAllResponse.PropertyTags[i].PropertyId.GetType(),
                    "MS-OXCDATA",
                    182,
                    @"[In PropertyTag Structure] PropertyId (2 bytes): An unsigned integer that identifies the property.");

                // PtypInteger32 Property Type Value is 0x0003.
                if (queryColumnsAllResponse.PropertyTags[i].PropertyType == 0x0003)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2691");

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2691
                    Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2691,
                        @"[In Property Data Types] PtypInteger32 (PT_LONG, PT_I4, int, ui4) is that 4 bytes; a 32-bit integer [MS-DTYP]: INT32 with Property Type Value 0x0003,%x03.00.");
                }

                // PtypInteger64 Property Type Value is 0x0014.
                if (queryColumnsAllResponse.PropertyTags[i].PropertyType == 0x0014)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2699");

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2699
                    Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2699,
                        @"[In Property Data Types] PtypInteger64 (PT_LONGLONG, PT_I8, i8, ui8) is that 8 bytes; a 64-bit integer [MS-DTYP]: LONGLONG with Property Type Value 0x0014,%x14.00.");
                }
            }
        }

        /// <summary>
        /// Verify RopFindRow Response
        /// </summary>
        /// <param name="findRowResponse">RopFindRowResponse structure data that needs verification</param>
        private void VerifyRopFindRowResponse(RopFindRowResponse findRowResponse)
        {
            this.VerifyRPCLayerRequirement();

            if (findRowResponse.ReturnValue == 0x00000000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R237");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R237
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    findRowResponse.RowNoLongerVisible.GetType(),
                    237,
                    @"[In RopFindRow ROP Response Buffer] RowNoLongerVisible (1 byte): A Boolean that indicates whether the row to which the bookmark pointed is no longer visible.");

                if (findRowResponse.HasRowData == 0x01)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R245");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R245
                    Site.CaptureRequirementIfIsNotNull(
                         findRowResponse.RowData,
                         245,
                        @"[In RopFindRow ROP Response Buffer] If the value of the HasRowData field is ""TRUE"" (0x01), the RowData field MUST be present.");

                    this.VerifyPropertyRowStructure(findRowResponse.RowData);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R244");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R244
                    Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(PropertyRow),
                        findRowResponse.RowData.GetType(),
                        244,
                        @"[In RopFindRow ROP Response Buffer] RowData (variable): A PropertyRow structure, as specified in [MS-OXCDATA] section 2.8.1, that specifies the row.");
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R242");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R242
                Site.CaptureRequirementIfIsTrue(
                    findRowResponse.HasRowData == 0x01 || findRowResponse.HasRowData == 0x00,
                    242,
                    @"[In RopFindRow ROP Response Buffer] HasRowData (1 byte): A Boolean that specifies whether a row is included in this response.");
            }

            if (findRowResponse.HasRowData == 0x00)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R246");

                // Verify MS-OXCTABL requirement: MS-OXCTABL_R246
                Site.CaptureRequirementIfIsNull(
                findRowResponse.RowData,
                246,
                @"[In RopFindRow ROP Response Buffer] If the value of HasRowData is ""FALSE"" (0x00), the RowData field MUST NOT be present.");
            }
        }

        /// <summary>
        /// Verify ropExpandRow
        /// </summary>
        /// <param name="expandRowResponse">RopExpandRowResponse structure data that needs verification</param>
        /// <param name="theValueMaxRowCountInExpandRowrequest">The value of the MaxRowCount field in the ROP request buffer</param>
        private void VerifyRopExpandRowResponse(RopExpandRowResponse expandRowResponse, ushort theValueMaxRowCountInExpandRowrequest)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R278");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R278
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                expandRowResponse.RowCount.GetType(),
                278,
                @"[In RopExpandRow ROP Response Buffer] ExpandedRowCount (4 bytes): An unsigned integer that specifies the total number of rows that are in the expanded category.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R279");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R279
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                expandRowResponse.RowCount.GetType(),
                279,
                @"[In RopExpandRow ROP Response Buffer] RowCount (2 bytes): An unsigned integer that specifies the number of PropertyRow structures, as specified in [MS-OXCDATA] section 2.8.1, that are contained in the RowData field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R280: the value of the RowCount field in the response for RopExpandRow is set to {0}", expandRowResponse.RowCount);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R280
            bool isVerifyR280 = expandRowResponse.RowCount <= theValueMaxRowCountInExpandRowrequest;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR280,
                280,
                @"[In RopExpandRow ROP Response Buffer] The value of this field [RowCount] MUST be less than or equal to the value of the MaxRowCount field in the ROP request buffer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R281: the value of the RowCount field is:{0}, and ExpandedRowCount is:{1}", expandRowResponse.RowCount, expandRowResponse.ExpandedRowCount);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R281
            bool isVerifyR281 = expandRowResponse.RowCount <= expandRowResponse.ExpandedRowCount;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR281,
                281,
                @"[In RopExpandRow ROP Response Buffer] The value of this field [RowCount] MUST be less than or equal to the value of the ExpandedRowCount field in the ROP response buffer.");
        }

        /// <summary>
        /// Verify RopCollapseRow Response
        /// </summary>
        /// <param name="collapseRowResponse">RopCollapseRowResponse structure data that needs verification</param>
        private void VerifyRopCollapseRowResponse(RopCollapseRowResponse collapseRowResponse)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R287 ");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R287 
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                collapseRowResponse.CollapsedRowCount.GetType(),
                287,
                @"[In RopCollapseRow ROP Response Buffer] CollapsedRowCount (4 bytes): An unsigned integer that specifies the number of rows that have been collapsed.");
        }

        /// <summary>
        /// Verify RopGetCollapseState Response
        /// </summary>
        /// <param name="getCollapseStateResponse">RopGetCollapseStateResponse structure data that needs verification</param>
        private void VerifyRopGetCollapseStateResponse(RopGetCollapseStateResponse getCollapseStateResponse)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R294: the value of the CollapseStateSize field is:{0},and the length of CollapseState field is:{1} in response", getCollapseStateResponse.CollapseStateSize, getCollapseStateResponse.CollapseState.Length);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R294
            bool isVerifyR294 = getCollapseStateResponse.CollapseStateSize == getCollapseStateResponse.CollapseState.Length;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR294,
                294,
                @"[In RopGetCollapseState ROP Response Buffer] CollapseStateSize (2 bytes): An unsigned integer that specifies the size, in bytes, of the CollapseState field.");
        }

        /// <summary>
        /// Verify RopGetCollapseState Response
        /// </summary>
        /// <param name="setCollapseStateResponse">RopGetCollapseStateResponse structure data that needs verification</param>
        private void VerifyRopSetCollapseStateResponse(RopSetCollapseStateResponse setCollapseStateResponse)
        {
            this.VerifyRPCLayerRequirement();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R314: the value of the BookmarkSize field is:{0},and the length of Bookmark field is:{1} in response", setCollapseStateResponse.BookmarkSize, setCollapseStateResponse.Bookmark.Length);

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R314
            bool isVerifyR314 = setCollapseStateResponse.BookmarkSize == setCollapseStateResponse.Bookmark.Length;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR314,
                314,
                @"[In RopSetCollapseState ROP Response Buffer] BookmarkSize (2 bytes): An unsigned integer that specifies the size, in bytes, of the Bookmark field.");
        }

        /// <summary>
        /// Verify PropertyRow Structure
        /// </summary>
        /// <param name="propertyRow">PropertyRow Structure returned by the RopFindRow or RopQueryRows ROPs</param>
        private void VerifyPropertyRowStructure(PropertyRow propertyRow)
        {
            // Since the returned PropertyRow Structure can be parsed correctly, the following requirement can be verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R65");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R65
            Site.CaptureRequirement(
                "MS-OXCDATA",
                65,
                @"[In PropertyRow Structures] For the RopFindRow, RopGetReceiveFolderTable, and RopQueryRows ROPs, property values are returned in the order of the properties in the table, set by a prior call to a RopSetColumns ROP request ([MS-OXCROPS] section 2.2.5.1).");

            // The flag with value 0x00 means the structure is StandardPropertyRow, and 0x01 means the structure is FlaggedPropertyRow Structure
            if (propertyRow.Flag == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R72");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R72
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    propertyRow.Flag.GetType(),
                    "MS-OXCDATA",
                    72,
                    @"[In StandardPropertyRow Structure] Flag (1 byte): An unsigned integer."); 

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R79");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R79
                Site.CaptureRequirement(
                    "MS-OXCDATA",
                    79,
                    @"[In FlaggedPropertyRow Structure] Flag (1 byte): Otherwise [when PtypUnspecified was not used in the ROP request and the ROP response includes a type], this value MUST be set to 0x00.");

                int i = 0;
                uint value = 0;
                bool hasErrorPropertyValue = false;
                for (; i < propertyRow.PropertyValues.Count; i++)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R75: the structure type of the property is {0}", propertyRow.PropertyValues[i].GetType());
                    bool isVerifyR75 = propertyRow.PropertyValues[i].GetType() == typeof(PropertyValue) || propertyRow.PropertyValues[i].GetType() == typeof(PropertyValue) || propertyRow.PropertyValues[i].GetType() == typeof(TypedPropertyValue);

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R75
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR75,
                        "MS-OXCDATA",
                        75,
                        @"[In StandardPropertyRow Structure] ValueArray (variable): At each position of the array, the structure will either be a PropertyValue structure, as specified in section 2.11.2.1, "
                        + "if the type of the corresponding property tag was specified, or a TypedPropertyValue structure, as specified in section 2.11.3, if the type of the corresponding property tag was PtypUnspecified (section 2.11.1).");

                    if (propertyRow.PropertyValues[i].Value.Length == 4)
                    {
                        value = BitConverter.ToUInt32(propertyRow.PropertyValues[i].Value, 0);
                        if (value == 0x8004010F || value == 0x80040302 || value == 0x80040303 || value == 0x80040304)
                        {
                            hasErrorPropertyValue = true;
                            break;
                        }
                    }

                    value = 0;
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R73{0}", hasErrorPropertyValue ? string.Format(": No.{0} property value is not correct and the value is {1}", i, value) : string.Empty);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R73
                bool isVerifyR73 = hasErrorPropertyValue;
                Site.CaptureRequirementIfIsFalse(
                    isVerifyR73,
                    "MS-OXCDATA",
                    73,
                    @"[In StandardPropertyRow Structure] Flag (1 byte): This value MUST be set to 0x00 to indicate that all property values are present and without error.");
            }
            else if (propertyRow.Flag == 1)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R76");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R76
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    propertyRow.Flag.GetType(),
                    "MS-OXCDATA",
                    76,
                    @"[In FlaggedPropertyRow Structure] Flag (1 byte): An unsigned integer."); 

                int i = 0;
                uint value = 0;
                bool hasErrorPropertyValue = false;
                for (; i < propertyRow.PropertyValues.Count; i++)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R81: the structure type of the property is {0}", propertyRow.PropertyValues[i].GetType());
                    bool isVerifyR81 = propertyRow.PropertyValues[i].GetType() == typeof(FlaggedPropertyValue) || propertyRow.PropertyValues[i].GetType() == typeof(FlaggedPropertyValueWithType);
                    
                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R81
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR81,
                        "MS-OXCDATA",
                        81,
                        @"[In FlaggedPropertyRow Structure] ValueArray (variable): At each position of the array, the structure will be either a FlaggedPropertyValue structure, as specified in section 2.11.5, "
                        + "if the type of the corresponding property tag was previously specified or a FlaggedPropertyValueWithType structure, as specified in section 2.11.6, if the type of the corresponding property tag was PtypUnspecified.");

                    if (propertyRow.PropertyValues[i].Value.Length == 4)
                    {
                        value = BitConverter.ToUInt32(propertyRow.PropertyValues[i].Value, 0);
                        if (value == 0x8004010F || value == 0x80040302 || value == 0x80040303 || value == 0x80040304)
                        {
                            hasErrorPropertyValue = true;
                            break;
                        }
                    }

                    value = 0;
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R77{0}", hasErrorPropertyValue ? string.Empty : "There is no property which contains error value!");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R77
                bool isVerifyR77 = hasErrorPropertyValue;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR77,
                    "MS-OXCDATA",
                    77,
                    @"[In FlaggedPropertyRow Structure] Flag (1 byte): This value [Flag] MUST be set to 0x01 to indicate that there are errors or some property values are missing.");
            }
            else
            {
                Site.Assert.Fail("The value of the flag field is expected to be set to 0x00 or 0x01, but the current value is {0}", propertyRow.Flag);
            }
        }

        /// <summary>
        /// Verify if the asynchronous ROP is complete
        /// </summary>
        /// <param name="getStatusResponse">RopGetStatusResponse structure data that needs verification</param>
        private void VerifyAsynchronousROPComplete(RopGetStatusResponse getStatusResponse)
        {
            if (Common.IsRequirementEnabled(891, this.Site))
            {
                if (this.globalIsRestrictAsynchronous || this.globalIsSortTableAsynchronous || this.globalIsSetColumnsAsynchronous)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R891");

                    // Verify MS-OXCTABL requirement: MS-OXCTABL_R891
                    // TableStatus in response with value 0x00 means the ROP is finished  
                    Site.CaptureRequirementIfAreEqual<byte>(
                        0x00,
                        getStatusResponse.TableStatus,
                        891,
                        @"[In Appendix A: Product Behavior] The RopGetStatus ROP of the implementation ([MS-OXCROPS] section 2.2.5.6) MUST send TBLSTAT_COMPLETE when the current asynchronous was finished on the table in the response, as specified in section 3.2.5.1. (Microsoft Exchange Server 2007 follows this behavior.)");
                }
            }
        }

        /// <summary>
        /// Verify RPC layer requirement
        /// </summary>
        private void VerifyRPCLayerRequirement()
        {
            // Since the request and response can be parsed correctly, the following requirement can be captured directly.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCTABL_R3");

            // Verify MS-OXCTABL requirement: MS-OXCTABL_R3
            Site.CaptureRequirement(
                3,
                @"[In Transport] The ROP request buffers and ROP response buffers specified by this protocol are sent to and received by the server by using the underlying Remote Operations (ROP) List and Encoding Protocol, as specified in [MS-OXCROPS].");
        }
    }
}