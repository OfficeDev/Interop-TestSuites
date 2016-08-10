[assembly: Microsoft.Xrt.Runtime.NativeType("System.Diagnostics.Tracing.*")]

namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    using System.Collections.Generic;
    using Microsoft.Modeling;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// MS-OXCTABL model program
    /// </summary>
    public static partial class Model
    {
        #region Global variables
        /// <summary>
        /// Identify the table type
        /// </summary>
        private static TableType globleTableType;

        /// <summary>
        /// Set columns flags
        /// </summary>
        private static Dictionary<uint, bool> setColumnsFlags = new Dictionary<uint, bool>();

        /// <summary>
        /// Sort table flags
        /// </summary>
        private static Dictionary<uint, bool> sortTableFlags = new Dictionary<uint, bool>();

        /// <summary>
        /// Restrict flags
        /// </summary>
        private static Dictionary<uint, bool> restrictFlags = new Dictionary<uint, bool>();

        /// <summary>
        /// Identify whether the bookmark is valid
        /// </summary>
        private static bool validBookmark = false;

        /// <summary>
        /// Identify whether a bookmark is created
        /// </summary>
        private static bool bookmarkCreated = false;

        /// <summary>
        /// Identify whether the table is reset
        /// </summary>
        private static bool resetTableDone = false;

        /// <summary>
        /// Identify whether the free bookmark is done
        /// </summary>
        private static bool freeBookmarkDone = false;

        /// <summary>
        /// Identify whether the RopRelease is done
        /// </summary>
        private static bool ropReleaseDone = false;

        /// <summary>
        /// Identify the requirement Id and the enable property value related to the optional behavior
        /// </summary>
        private static MapContainer<int, bool> requirementContainer = new MapContainer<int, bool>();
        #endregion

        #region Rule methods

        #region Initial methods
        /// <summary>
        /// Determines if the requirement is enabled or not.
        /// </summary>
        /// <param name="requirementId">The requirement Id</param>
        /// <param name="enabled">Requirement is enable or not.</param>
        [Rule(Action = "CheckRequirementEnabled(requirementId,out enabled)")]
        public static void CheckRequirementEnabled(int requirementId, out bool enabled)
        {
            enabled = Choice.Some<bool>();
            requirementContainer.Add(requirementId, enabled);
        }

        /// <summary>
        /// This method is used to check whether MAPIHTTP transport is supported by test suite.
        /// </summary>
        /// <param name="isSupported">The transport is supported or not.</param>
        [Rule(Action = "CheckMAPIHTTPTransportSupported(out isSupported)")]
        public static void CheckMAPIHTTPTransportSupported(out bool isSupported)
        {
            isSupported = Choice.Some<bool>();
        }

        /// <summary>
        /// This method is used to initialize the table
        /// </summary>
        /// <param name="tableType">Identify the table type</param>
        [Rule(Action = "InitializeTable(tableType)")]
        public static void InitializeTable(TableType tableType)
        {
            // Initialize the table type
            globleTableType = tableType;
            setColumnsFlags.Clear();
        }

        /// <summary>
        /// This method is used to disconnect from server
        /// </summary>
        [Rule(Action = "Disconnect()")]
        public static void Disconnect()
        {

        }

        #endregion

        #region RopSetColumns
        /// <summary>
        /// This method is used to synchronize SetColumns for a table
        /// </summary>
        /// <param name="reqId">Identify the request ID</param>
        /// <param name="invalidPropertyTag">Indicate whether there are invalid property tags in the column set </param>
        /// <param name="isNewColumnSet">Indicate whether to change the current column set with new one</param>
        /// <param name="isSetColumnsAsynchronous">Indicate whether RopsetColumns Rop is to be performed asynchronously</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopSetColumns(reqId,invalidPropertyTag,isNewColumnSet,isSetColumnsAsynchronous)/result")]
        public static TableRopReturnValues RopSetColumns(uint reqId, bool invalidPropertyTag, bool isNewColumnSet, bool isSetColumnsAsynchronous)
        {
            // This ROP is valid only on table objects.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                // The RopSetColumns ROP MUST be supported for contents tables.
                ModelHelper.CaptureRequirement(831, @"[In Processing RopSetColumns] The RopSetColumns ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // The RopSetColumns ROP MUST be supported for hierarchy tables.
                ModelHelper.CaptureRequirement(829, @"[In Processing RopSetColumns] The RopSetColumns ROP MUST be supported for hierarchy tables.");
            }
            else if (globleTableType == TableType.RULES_TABLE)
            {
                // The RopSetColumns ROP MUST be supported for attachments tables.
                ModelHelper.CaptureRequirement(830, @"[In Processing RopSetColumns] The RopSetColumns ROP MUST be supported for rules tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned. In this test suite, invalid table is referred as a folderId. 
                // For details, see the table initial method in adapter project.
                ModelHelper.CaptureRequirement(436, @"[In Processing RopSetColumns] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not of type table.");

                // Record the setColumns flags.
                setColumnsFlags.Add(reqId, false);

                // The value ecNotSupported will be returned if the object on which this ROP was sent is not of type table.
                return TableRopReturnValues.ecNotSupported;
            }

            // Invalid Property Tag case
            if (invalidPropertyTag)
            {
                ModelHelper.CaptureRequirement(
                    435,
                    @"[In Processing RopSetColumns] The error code ecInvalidParam will be returned with value 0x80070057,%x57.00.07.80 means a property tag in the column array is of type PT_UNSPECIFIED, PT_ERROR, or an invalid type.");

                // Record the setColumns flags 
                setColumnsFlags.Add(reqId, false);

                // ecInvalidParam will be returned if a property tag in the column array is of type
                // when invalidPropertyTag is set to true, PT_UNSPECIFIED will be include in the column set, it does not need to set other invalid property tag because all these properties can be implied by PT_UNSPECIFIED, 
                return TableRopReturnValues.ecInvalidParam;
            }

            // Record the setColumns flags 
            setColumnsFlags.Add(reqId, true);

            // RopSetColumns is valid only on Table objects.
            ModelHelper.CaptureRequirement(45, @"[In RopSetColumns ROP] This ROP [RopSetColumns] is valid only on Table objects.");
            return TableRopReturnValues.success;
        }
        #endregion

        #region RopSortTable
        /// <summary>
        /// This method is used to sort table synchronously
        /// </summary>
        /// <param name="reqId">Identify the request ID</param>
        /// <param name="validHandle">Identify whether the Handle is valid, this is used to trigger the error code in SortTable</param>
        /// <param name="multipleSortOrders">Indicate whether sort orders contains more than one sort order structure</param>
        /// <param name="isMaximumCategory">Indicate whether sort orders contains a sort order structure with the order set to MaximumCategory</param>
        /// <param name="allSortOrdersUsedAsCategory">Identify whether all sort orders used as category</param>
        /// <param name="allCategoryExpanded">Identify whether all categories are expanded</param>
        /// <param name="isNewOrder">Identify whether to change the current sort order</param>
        /// <param name="isSortTableAsynchronous">Indicate whether RopSortTable Rop is to be performed asynchronously</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopSortTable(reqId,validHandle,multipleSortOrders,isMaximumCategory,allSortOrdersUsedAsCategory,allCategoryExpanded,isNewOrder,isSortTableAsynchronous)/result")]
        public static TableRopReturnValues RopSortTable(uint reqId, bool validHandle, bool multipleSortOrders, bool isMaximumCategory, bool allSortOrdersUsedAsCategory, bool allCategoryExpanded, bool isNewOrder, bool isSortTableAsynchronous)
        {
            // This ROP is valid only on contents table.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                ModelHelper.CaptureRequirement(447, @"[In Processing RopSortTable] The RopSortTable ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE || globleTableType == TableType.RULES_TABLE || globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                // Record the sortTableFlags
                sortTableFlags.Add(reqId, false);

                // If the table type is not content table, ecNotSupported will be returned.
                ModelHelper.CaptureRequirement(450, @"[In Processing RopSortTable] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not a contents table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else
            {
                // Record the sortTableFlags
                sortTableFlags.Add(reqId, false);
                ModelHelper.CaptureRequirement(57, @"[In RopSortTable ROP] This ROP is valid only on Table objects.");

                // If invalid table, ecNotSupported will be returned. in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project   
                return TableRopReturnValues.ecNotSupported;
            }

            // If validHandle is false, then ecNullObject will be returned, this code block is used to trigger 
            // a failure case for sort table.
            if (!validHandle)
            {
                sortTableFlags.Add(reqId, false);
                return TableRopReturnValues.ecNullObject;
            }

            // Record the sortTableFlags
            sortTableFlags.Add(reqId, true);

            // If a bookmark has been created, invalid it. 
            // 3.2.5.3,"When this ROP is sent, the server MUST invalidate all current bookmarks of the table"
            if (validBookmark)
            {
                validBookmark = false;
            }

            return TableRopReturnValues.success;
        }
        #endregion

        #region RopRestrict
        /// <summary>
        /// This method is used to Restrict for a table synchronously
        /// </summary>
        /// <param name="reqId">Identify the request ID</param>
        /// <param name="validLogonId">Identify whether the logonId is valid, this is used to trigger the error code in Restrict</param>
        /// <param name="isNewRestrict">A flag identify whether new restriction is required</param>
        /// <param name="isRestrictAsynchronous">Indicate whether RopRestrict Rop is to be performed asynchronously</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopRestrict(reqId,validLogonId,isNewRestrict,isRestrictAsynchronous)/result")]
        public static TableRopReturnValues RopRestrict(uint reqId, bool validLogonId, bool isNewRestrict, bool isRestrictAsynchronous)
        {
            // The RopRestrict ROP is  valid only on contents table, hierarchy table, and rules table.
            if (globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                restrictFlags.Add(reqId, false);

                // The error code ecNotSupported will be returned with value 0x80040102(%x02.01.04.80),
                // if the object on which this ROP was sent is not a contents or hierarchy table.");
                ModelHelper.CaptureRequirement(461, @"[In Processing RopRestrict] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not a contents table, hierarchy table, or rules table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else if (globleTableType == TableType.CONTENT_TABLE)
            {
                // RopRestrict MUST be supported for contents tables.
                ModelHelper.CaptureRequirement(834, @"[In Processing RopRestrict] The RopRestrict ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // RopRestrict MUST be supported for hierarchy tables.
                ModelHelper.CaptureRequirement(833, @"[In Processing RopRestrict] The RopRestrict ROP MUST be supported for hierarchy tables.");
            }
            else if (globleTableType == TableType.RULES_TABLE)
            {
                // RopRestrict MUST be supported for rules tables.
                ModelHelper.CaptureRequirement(832, @"[In Processing RopRestrict] The RopRestrict ROP MUST be supported for rules tables.");
            }
            else
            {
                restrictFlags.Add(reqId, false);

                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(88, @"[In RopRestrict ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            // If validLogon is false, then ecNullObject will be returned, this code block is used to trigger a failure case for restrict table.
            if (!validLogonId)
            {
                restrictFlags.Add(reqId, false);
                return TableRopReturnValues.ecNullObject;
            }

            // Record the restrictFlags.
            restrictFlags.Add(reqId, true);

            // If a bookmark has been created, invalid it.
            // Record the signal, for the detailed requirement capture coded, it is used to verify 453 in RopQueryPositionResponse Action, 
            // see capture logic for 453.
            if (validBookmark)
            {
                validBookmark = false;
            }

            return TableRopReturnValues.success;
        }
        #endregion

        #region RopQueryRows
        /// <summary>
        /// This method is used to query rows for a table
        /// </summary>
        /// <param name="queryRowFlags">A flag for the query row operation</param>
        /// <param name="isForwardRead">Identify whether to read the data forward</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopQueryRows(queryRowFlags,isForwardRead)/result")]
        public static TableRopReturnValues RopQueryRows(QueryRowsFlags queryRowFlags, bool isForwardRead)
        {
            // This ROP is only valid on Table objects.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                // The RopQueryRows ROP MUST be supported for contents tables.
                ModelHelper.CaptureRequirement(837, @"[In Processing RopQueryRows] The RopQueryRows ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // The RopQueryRows ROP MUST be supported for hierarchy tables.
                ModelHelper.CaptureRequirement(835, @"[In Processing RopQueryRows] The RopQueryRows ROP MUST be supported for hierarchy tables.");
            }
            else if (globleTableType == TableType.RULES_TABLE)
            {
                // The RopQueryRows ROP MUST be supported for attachments tables.
                ModelHelper.CaptureRequirement(836, @"[In Processing RopQueryRows] The RopQueryRows ROP MUST be supported for rules tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(477, @"[In Processing RopQueryRows] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not of type table.");
                ModelHelper.CaptureRequirement(614, @"[In RopQueryRows ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            // Test RopSetColumns MUST be sent on the table prior to sending RopQueryRows
            bool latestSetColumnFlag = false;
            setColumnsFlags.TryGetValue((uint)setColumnsFlags.Count, out latestSetColumnFlag);

            // If there is no successful SetColumn response
            if (!latestSetColumnFlag)
            {
                Condition.IsTrue((requirementContainer[768] && requirementContainer[866] && requirementContainer[867]) ||
                    (!requirementContainer[768] && !requirementContainer[866] && !requirementContainer[867]));
                if (requirementContainer[768] && requirementContainer[866] && requirementContainer[867])
                {
                    ModelHelper.CaptureRequirement(768, @"[In Appendix A: Product Behavior] If a RopQueryRows ROP is sent before a successful RopSetColumns ROP, then the implementation fails the ROP with ""ecNullObject"". (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    ModelHelper.CaptureRequirement(866, @"[In Appendix A: Product Behavior] The RopSetColumns ROP ([MS-OXCROPS] section 2.2.5.1) MUST be sent on the table before sending a RopQueryRows ROP request on the implementation. (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    ModelHelper.CaptureRequirement(867, @"[In Appendix A: Product Behavior] The error code ecNullObject will be returned with value 0x000004B9,%xB9.04.00.00, if RopSetColumns has not been sent on this table on the implementation. (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    return TableRopReturnValues.ecNullObject;
                }
                else
                {
                    if ((uint)setColumnsFlags.Count == 0)
                    {
                        return TableRopReturnValues.unexpected;
                    }
                }
            }

            return TableRopReturnValues.success;
        }

        /// <summary>
        /// This method is used to handle value which RopQueryRowsRequest returned
        /// </summary>
        /// <param name="queryRowFlags">A flag for the query row operation</param>
        /// <param name="isForwardRead">Identify whether to read the data forward</param>
        /// <param name="isZeroRow">Identify whether zero row returned</param>
        /// <param name="ropType">Identify ROP type</param>
        /// <param name="isCorrectRowCount">Identify whether the correct row count data was read</param>
        /// <param name="isCursorPositionChanged">Identify whether the cursor position was changed</param>
        /// <param name="isLatestRopData">Identify whether the queried data equals to the last ROP data</param>
        /// <param name="isLastSuccessRopData">Identify whether the queried row data equals to the last successful ROP data</param>
        /// <param name="queryRowOrigin">Identify the cursor position returned in queryRow response</param>
        /// <param name="isRequestCountTooLarger">Identify whether the row count in the request is larger than the actual row count</param>
        [Rule(Action = "RopQueryRowsResponse(queryRowFlags,isForwardRead,isZeroRow,ropType,isCorrectRowCount,isCursorPositionChanged,isLatestRopData,isLastSuccessRopData,queryRowOrigin,isRequestCountTooLarger)")]
        public static void RopQueryRowsResponse(QueryRowsFlags queryRowFlags, bool isForwardRead, bool isZeroRow, TableRopType ropType, bool isCorrectRowCount, bool isCursorPositionChanged, bool isLatestRopData, bool isLastSuccessRopData, BookmarkType queryRowOrigin, bool isRequestCountTooLarger)
        {
            // If the NoAdvance flag (0x01) is set in the QueryRowsFlags field, the server MUST NOT change the position of the cursor.
            if (queryRowFlags == QueryRowsFlags.NoAdvance)
            {
                Condition.IsTrue(!isCursorPositionChanged);
                ModelHelper.CaptureRequirement(470, @"[In Processing RopQueryRows] If the NoAdvance flag (0x01) is set in the QueryRowsFlags field, the server MUST NOT change the position of the cursor.");
                if (queryRowOrigin == BookmarkType.BOOKMARK_BEGINNING)
                {
                    ModelHelper.CaptureRequirement(27, @"[In Predefined Bookmarks] When the Predefined Bookmarks is BOOKMARK_BEGINNING with value 0x00 means points to the beginning position of the table, or the first row.");
                }
            }

            // Test RopSetColumns MUST be sent on the table prior to sending RopQueryRows
            bool latestSetColumnFlag = false;
            setColumnsFlags.TryGetValue((uint)setColumnsFlags.Count, out latestSetColumnFlag);
            if (latestSetColumnFlag)
            {
                Condition.IsTrue(isCorrectRowCount);

                // If the ForwardRead field is TRUE, RopQueryRows MUST return the rows beginning at the origin, reading forward.
                if (isForwardRead)
                {
                    // In the Test Suite, the origin is beginning, and if the enough row is true, 
                    // the RowCount rows will be queried from table. Also, here it only verify whether the correct row count 
                    // is returned, because for the details of the table row, it can't be verified.
                    ModelHelper.CaptureRequirement(468, @"[In Processing RopQueryRows] If rows are returned by the RopQueryRows ROP and the ForwardRead field is set to ""TRUE"" (0x01), the RopQueryRows ROP returns the rows beginning at the position pointed to by the Origin field, reading forward.");
                }
                else
                { 
                    Condition.IsTrue((requirementContainer[896] && !requirementContainer[897]) || (!requirementContainer[896] && requirementContainer[897]));

                    // if there are no more rows to return and the ForwardRead field in the ROP request is set to "false",
                    // Exchange 2007 will set the Origin field to "BOOKMARK_BEGINNING".
                    if (requirementContainer[896] && isRequestCountTooLarger)
                    {
                        Condition.IsTrue(queryRowOrigin == BookmarkType.BOOKMARK_BEGINNING);
                        ModelHelper.CaptureRequirement(896, @"[In Appendix A: Product Behavior] Implementation does set the Origin field to ""BOOKMARK_BEGINNING"", if there are no more rows to return and the ForwardRead field in the ROP request is set to ""FALSE"". (Exchange Server 2007 follows this behavior.)");
                    }

                    // if there are no more rows to return and the ForwardRead field in the ROP request is set to "false",
                    // Exchange 2010 and above will set the Origin field to "BOOKMARK_END".
                    if (requirementContainer[897] && isRequestCountTooLarger)
                    {
                        Condition.IsTrue(queryRowOrigin == BookmarkType.BOOKMARK_END);
                        ModelHelper.CaptureRequirement(897, @"[In Appendix A: Product Behavior] Implementation does set the Origin field to ""BOOKMARK_END"", if there are no more rows to return and the ForwardRead field in the ROP request is set to ""FALSE"". (<25> Section 3.2.5.5: Exchange 2010, Exchange 2013, and Exchange 2016 set the Origin field to ""BOOKMARK_END"".)");
                    }

                    // When ForwardRead in request is false, if the correct row count is returned, this requirement can be verified.
                    ModelHelper.CaptureRequirement(469, @"[In Processing RopQueryRows] If it [ForeardRead] is set to ""FALSE"" (0x00), the server returns the rows starting at the number of rows equal to the value of the RowCount field before the position pointed to by the Origin field, such that the rows returned are those between the value of the Origin field minus the value of the RowCount field and the position pointed to by the Origin field.");
                }

                // If the RopRestrict ROP has not been issued, the RowData field MUST include all rows.
                bool isLastestRestrictFlags = false;
                restrictFlags.TryGetValue((uint)restrictFlags.Count, out isLastestRestrictFlags);
                
                if (!isLastestRestrictFlags)
                {
                    ModelHelper.CaptureRequirement(118, @"[In RopQueryRows ROP Response Buffer] If the RopRestrict ROP has not been issued, the RowData field MUST include all rows.");
                }

                // When correct row count is returned and the returned row count is zero, the R467 is verified. 
                if (isZeroRow)
                {
                    ModelHelper.CaptureRequirement(467, @"[In Processing RopQueryRows] This ROP MUST only send zero rows when there are no more rows in the table.");
                }
            }

            ModelHelper.CaptureRequirement(462, @"[In Processing RopQueryRows] When a RopQueryRows ROP request ([MS-OXCROPS] section 2.2.5.4) is sent, the server MUST send the rows from the table starting at the current cursor position.");

            // If there are no more rows to return, the Origin field will be set to BOOKMARK_END if the ForwardRead field in the ROP request is set to true.
            if (isCursorPositionChanged && isForwardRead && queryRowOrigin == BookmarkType.BOOKMARK_END)
            {
                ModelHelper.CaptureRequirement(29, @"[In Predefined Bookmarks] When the Predefined Bookmarks is BOOKMARK_END with value 0x02 means points to the ending position of the table, or the location after the last row.");
                ModelHelper.CaptureRequirement(725, @"[In Processing RopQueryRows] If there are no more rows to return and the ForwardRead field in the ROP request is set to ""TRUE"", then the Origin field is set to ""BOOKMARK_END"".");
            }

            if (!isRequestCountTooLarger && isForwardRead && isCursorPositionChanged)
            {
                Condition.IsTrue(queryRowOrigin == BookmarkType.BOOKMARK_CURRENT);
                ModelHelper.CaptureRequirement(28, @"[In Predefined Bookmarks] When the Predefined Bookmarks is BOOKMARK_CURRENT with value 0x01 means points to the current position of the table, or the current row.");
                ModelHelper.CaptureRequirement(726, @"[In Processing RopQueryRows] Otherwise,[If there are more rows to return] it [Origin] will be set to ""BOOKMARK_CURRENT"".");
            }

            // Below is to parse the data queried
            switch (ropType)
            {
                // Query rows is called after a SetColumn request
                case TableRopType.SETCOLUMNS:
                    {
                        ParseQueryRowsAfterSetColumns(isLatestRopData, isLastSuccessRopData);
                        break;
                    }

                // Query rows is called after sort table request
                case TableRopType.SORTTABLE:
                    {
                        ParseQueryRowsAfterSortTable(isLatestRopData, isLastSuccessRopData);
                        break;
                    }

                // Query rows is called after restrict table request
                case TableRopType.RESTRICT:
                    {
                        ParseQueryRowsAfterRestrict(isLatestRopData, isLastSuccessRopData);
                        break;
                    }

                // If no above operation, query rows must return all rows, the parameter bIsLatestRopData is used to identify whether all row data is returned
                default:
                    {
                        break;
                    }
            }
        }
        #endregion

        #region RopQueryPosition
        /// <summary>
        /// This method is used to query position for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopQueryPosition()/result")]
        public static TableRopReturnValues RopQueryPosition()
        {
            // This ROP is only valid on Table objects.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                // The RopQueryPosition ROP MUST be supported for contents tables.
                ModelHelper.CaptureRequirement(846, @"[In Processing RopQueryPosition] The RopQueryPosition ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // The RopQueryPosition ROP MUST be supported for hierarchy tables.
                ModelHelper.CaptureRequirement(844, @"[In Processing RopQueryPosition] The RopQueryPosition ROP MUST be supported for hierarchy tables.");
            }
            else if (globleTableType == TableType.RULES_TABLE)
            {
                // The RopQueryPosition ROP MUST be supported for attachments tables.
                ModelHelper.CaptureRequirement(845, @"[In Processing RopQueryPosition] The RopQueryPosition ROP MUST be supported for rules tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(495, @"[In Processing RopQueryPosition] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not of type table.");
                ModelHelper.CaptureRequirement(612, @"[In RopQueryPosition ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            return TableRopReturnValues.success;
        }

        /// <summary>
        /// This method is used to handle value which RopQueryPositionRequest returned
        /// </summary>
        /// <param name="cursorPosition">Identify current cursor position</param>
        /// <param name="isRowCount">Identify whether the right rows count is returned</param>
        [Rule(Action = "RopQueryPositionResponse(cursorPosition,isRowCount)")]
        public static void RopQueryPositionResponse(CursorPosition cursorPosition, bool isRowCount)
        {
            // Cursor is current position and  total number of rows in response.
            Condition.IsTrue(isRowCount && (cursorPosition == CursorPosition.BEGIN
                    || cursorPosition == CursorPosition.CURRENT
                    || cursorPosition == CursorPosition.END));
            ModelHelper.CaptureRequirement(491, @"[In Processing RopQueryPosition] The RopQueryPosition ROP ([MS-OXCROPS] section 2.2.5.7) MUST send the current position of the cursor and the total number of rows in the table in the response.");

            // The server MUST invalidate all current bookmarks of the table and MUST move the cursor position to the beginning of the table.
            if (bookmarkCreated && restrictFlags.ContainsValue(true))
            {
                // After ROPSetColumns ROP, restrict, it need to use ROPQueryRows ROP to verify whether the columns, restrict set in the previous calls works. So this requirement is verified here.
                Condition.IsTrue(!validBookmark && (cursorPosition == CursorPosition.BEGIN)); 
                ModelHelper.CaptureRequirement(453, "[In Processing RopRestrict] When this ROP is sent, the server MUST invalidate all current bookmarks of the table and MUST move the cursor position to the beginning of the table.");
            }
        }
        #endregion

        #region RopSeekRow
        /// <summary>
        /// This method is used to seek row for a table
        /// </summary>
        /// <param name="origin">Identify the origin position</param>
        /// <param name="isForwardSeek">Identify whether to seek row forward</param>
        /// <param name="isWantRowMovedCountout">Identify whether to return the actual moved row number</param>
        /// <param name="isEnoughRow">Identify whether there is enough row to seek</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopSeekRow(origin,isForwardSeek,isWantRowMovedCountout,isEnoughRow)/result")]
        public static TableRopReturnValues RopSeekRow(BookmarkType origin, bool isForwardSeek, bool isWantRowMovedCountout, bool isEnoughRow)
        {
            // This ROP is only valid on Table objects.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                // The RopSeekRow ROP MUST be supported for contents tables.
                ModelHelper.CaptureRequirement(850, @"[In Processing RopSeekRow] The RopSeekRow ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // The RopSeekRow ROP MUST be supported for hierarchy tables.
                ModelHelper.CaptureRequirement(848, @"[In Processing RopSeekRow] The RopSeekRow ROP MUST be supported for hierarchy tables.");
            }
            else if (globleTableType == TableType.RULES_TABLE)
            {
                // The RopSeekRow ROP MUST be supported for attachments tables.
                ModelHelper.CaptureRequirement(849, @"[In Processing RopSeekRow] The RopSeekRow ROP MUST be supported for rules tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(142, @"[In RopSeekRow ROP] This ROP is valid only on Table objects.");
                ModelHelper.CaptureRequirement(502, @"[In Processing RopSeekRow] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80 ,if the object on which this ROP was sent is not of type table.");
                return TableRopReturnValues.ecNotSupported;
            }

            return TableRopReturnValues.success;
        }

        /// <summary>
        /// This method is used to handle value which RopSeekRowRequest returned
        /// </summary>
        /// <param name="isWantRowMovedCountout">Identify whether to return the actual seek row count</param>
        /// <param name="isEnoughRow">Identify whether there are enough row to seek</param>
        /// <param name="hasSoughtLess">Identify whether the actual seek row is less the one requested</param>
        /// <param name="isActualRowsCount">Identify whether the actual seek row is correct</param>
        [Rule(Action = "RopSeekRowResponse(isWantRowMovedCountout,isEnoughRow,hasSoughtLess,isActualRowsCount)")]
        public static void RopSeekRowResponse(bool isWantRowMovedCountout, bool isEnoughRow, bool hasSoughtLess, bool isActualRowsCount)
        {
            if (isWantRowMovedCountout)
            {
                if (!isEnoughRow)
                {
                    Condition.IsTrue(hasSoughtLess && isActualRowsCount);
                    ModelHelper.CaptureRequirement(154, @"[In RopSeekRow ROP Response Buffer] This field [HasSoughtLess] MUST be set to ""TRUE"" (0x01) if the number of rows moved is less than the number of rows requested (RowCount field).");
                    ModelHelper.CaptureRequirement(157, @"[In RopSeekRow ROP Response Buffer] This field's [HasSoughtLess'] value MUST be valid if the WantRowMovedCount field (in the request) is set to ""TRUE"".");
                
                    // RowsSought specifies the actual number of rows moved, so if the actualRowcount is true, MS-OXCTABL_R163 can be verified.
                    ModelHelper.CaptureRequirement(163, @"[In RopSeekRow ROP Response Buffer] This field's [RowsSought's] value MUST be valid if the WantRowMovedCount field (in the request) is set to ""TRUE"".");
                    ModelHelper.CaptureRequirement(150, @"[In RopSeekRow ROP Request Buffer] If this field [WantRowMovedCount] is set to ""TRUE"" (0x01), the server MUST return the actual number of rows moved.");
                    ModelHelper.CaptureRequirement(497, @"[In Processing RopSeekRow] If moving the cursor the number of rows equal to the value of the RowCount field would put it past the end (or beginning, if seeking backwards) of the table, and the WantRowMovedCount field is set to ""TRUE"" (0x01) in the request, the server MUST set the HasSoughtLess field to ""TRUE"" (0x01) and set the RowsSought field to the actual number of rows moved to reach the end of the table (or, in the case of seeking backwards, to the beginning of the table).");
                    ModelHelper.CaptureRequirement(152, @"[In RopSeekRow ROP Request Buffer] The actual number of rows moved can differ from the requested number of rows if the beginning or end of the table is encountered before moving the requested number of rows.");
                }
                else
                {
                    Condition.IsTrue(!hasSoughtLess && isActualRowsCount);
                    ModelHelper.CaptureRequirement(155, @"[In RopSeekRow ROP Response Buffer] Otherwise it [HasSoughtLess] MUST be set to ""FALSE"" (0x00) [ if the number of rows moved is not less than the number of rows requested (RowCount field)].");
                    ModelHelper.CaptureRequirement(496, @"[In Processing RopSeekRow] The RopSeekRow ROP ([MS-OXCROPS] section 2.2.5.8) MUST move the cursor position according to its request fields.");
                    
                    // In the request, it uses adapter to specify a pre-defined bookmark and the number of rows to move (forward or backwards), so, if the actualRowcount is true, MS-OXCTABL_R141 can be verified.
                    ModelHelper.CaptureRequirement(141, @"[In RopSeekRow ROP] The new location is specified by a predefined bookmark, as specified in section 2.2.2.1.1, and the number of rows to move (forward or backwards) from that bookmark.");
                }
            }
        }
        #endregion

        #region RopSeekRowBookmark
        /// <summary>
        /// This method is used to seek row bookmark for a table
        /// </summary>
        /// <param name="isForwardSeek">Identify whether to seek bookmark forward</param>
        /// <param name="isWantRowMovedCountout">Identify whether to return the actual moved row number</param>
        /// <param name="isEnoughRow">Identify whether there is enough row to seek</param>
        /// <returns>Table ROP Return Values</returns>
        [Rule(Action = "RopSeekRowBookmark(isForwardSeek,isWantRowMovedCountout,isEnoughRow)/result")]
        public static TableRopReturnValues RopSeekRowBookmark(bool isForwardSeek, bool isWantRowMovedCountout, bool isEnoughRow)
        {
            // This ROP is valid only on contents table and hierarchy table
            if (globleTableType == TableType.RULES_TABLE || globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                // The error code ecNotSupported will be returned with value 0x80040102 (%x02.01.04.80) means if the object on which this ROP was sent is not a contents or hierarchy table.
                ModelHelper.CaptureRequirement(512, @"[In Processing RopSeekRowBookmark] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80 means if the object on which this ROP was sent is not a contents or hierarchy table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else if (globleTableType == TableType.CONTENT_TABLE)
            {
                // The RopSeekRowBookmark ROP MUST be supported for contents tables.
                ModelHelper.CaptureRequirement(508, @"[In Processing RopSeekRowBookmark] The RopSeekRowBookmark ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // The RopSeekRowBookmark ROP MUST be supported for hierarchy tables.
                ModelHelper.CaptureRequirement(852, @"[In Processing RopSeekRowBookmark] The RopSeekRowBookmark ROP MUST be supported for hierarchy tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(168, @"[In RopSeekRowBookmark ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            // Before a success SetColumnsRop, RopSeekRowBookmark will fail in Exchange2010 and above
            if (!setColumnsFlags.ContainsValue(true) && !resetTableDone)
            {
                if (requirementContainer[8276])
                {
                    ModelHelper.CaptureRequirement(8276, @"[In Appendix A: Product Behavior] If a RopSeekRowBookmark ROP is sent before a successful RopSetColumns ROP, then the implementation fails the ROP with ""ecNullObject"". (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    return TableRopReturnValues.ecNullObject;
                }
                else
                {
                    return TableRopReturnValues.unexpected;
                }
            }

            if (ropReleaseDone)
            {
                ModelHelper.CaptureRequirement(248, @"[In RopFreeBookmark ROP] If the bookmark has been released by the ROPRelease ROP ([MS-OXCROPS] section 2.2.15.3), attempts to use the bookmark will fail with the ecNullObject error code.");
                ModelHelper.CaptureRequirement(552, @"[In Processing RopFreeBookmark] The error code ecNullObiect will be returned with value 0x000004B9 %xB9.04.00.00 means attempted to use the bookmark after it was released by the ROPRelease ROP ([MS-OXCROPS] section 2.2.15.3).");
                return TableRopReturnValues.ecNullObject;
            }

            if (!validBookmark)
            {
                if (freeBookmarkDone)
                {
                    // The error code ecInvalidBookmark will be returned if the bookmark (2) sent in the request is no longer valid.
                    ModelHelper.CaptureRequirement(511, "[In Processing RopSeekRowBookmark] The error code ecInvalidBookmark will be returned with value 0x80040405,%x05.04.04.80 means if the bookmark sent in the request is no longer valid.");
                    ModelHelper.CaptureRequirement(2481, "[In RopFreeBookmark ROP] If the bookmark is released by the RopFreeBookmark ROP, the server will return the ecInvalidBookmark error code when it is used.");
                    ModelHelper.CaptureRequirement(5521, "[In Processing RopFreeBookmark] The error code ecInvalidBookmark will be returned with value 0x80040405 %x05.04.04.80 means attempted to use the bookmark after it was released by the RopFreeBookmark ROP (section 2.2.2.15).");
                    return TableRopReturnValues.ecInvalidBookmark;
                }

                // After resetTable Rop, attempts to use the bookmark will return success in Exchange 2007.
                // After resetTable Rop, attempts to use the bookmark will return ecInvalidBookmark in Exchange 2010 and above
                if (resetTableDone)
                {
                    if (requirementContainer[909])
                    {
                        ModelHelper.CaptureRequirement(909, @"[In Appendix A: Product Behavior] Implementation returns ecSuccess [for RopSeekRowBookmark], if the bookmark has become invalid because of a RopResetTable ([MS-OXCROPS] section 2.2.5.15) ROP request. (<28> Section 3.2.5.10: Exchange 2007 returns ecSuccess.)");
                    }

                    if (requirementContainer[908])
                    {
                        ModelHelper.CaptureRequirement(908, @"[In Appendix A: Product Behavior] Implementation does set the ReturnValue field to ""ecInvalidBookmark"" [for RopSeekRowBookmark], if the bookmark has become invalid because of a RopResetTable ([MS-OXCROPS] section 2.2.5.15) ROP request. (Exchange Server 2010 and above follow this behavior.)");
                        return TableRopReturnValues.ecInvalidBookmark;
                    }
                }

                // After sortTable Rop, attempts to use the bookmark will fail with ecInvalidBookmark except Exchange 2007.
                if (sortTableFlags.ContainsValue(true))
                {
                    Condition.IsTrue((requirementContainer[904] && !requirementContainer[905]) || (!requirementContainer[904] && requirementContainer[905]));
                    if (requirementContainer[904])
                    {
                        ModelHelper.CaptureRequirement(904, @"[In Appendix A: Product Behavior] Implementation does set the ReturnValue field to ""ecInvalidBookmark"" [for RopSeekRowBookmark], if the bookmark has become invalid because of a RopSortTable ([MS-OXCROPS] section 2.2.5.2) ROP request. (Exchange Server 2010 and above follow this behavior.)");
                        return TableRopReturnValues.ecInvalidBookmark;
                    }

                    if (requirementContainer[905])
                    {
                        ModelHelper.CaptureRequirement(905, @"[In Appendix A: Product Behavior] Implementation returns ecSuccess [for RopSeekRowBookmark], if the bookmark has become invalid because of a RopSortTable ([MS-OXCROPS] section 2.2.5.2) ROP request. (<28> Section 3.2.5.10: Exchange 2007 returns ecSuccess.)");
                    }
                }

                // After restrict Table Rop, attempts to use the bookmark will fail with ecInvalidBookmark in Exchange 2010 and above.
                if (restrictFlags.ContainsValue(true))
                {
                    if (requirementContainer[906])
                    {
                        ModelHelper.CaptureRequirement(906, @"[In Appendix A: Product Behavior] Implementation does set the ReturnValue field to ""ecInvalidBookmark"" [for RopSeekRowBookmark], if the bookmark has become invalid because of a RopRestrict ([MS-OXCROPS] section 2.2.5.3) ROP request. (Exchange Server 2010 and above follow this behavior.)");
                        return TableRopReturnValues.ecInvalidBookmark;
                    }
                }
            }

            ModelHelper.CaptureRequirement(504, @"[In Processing RopSeekRowBookmark] It acts in the same way as the RopSeekRow ROP ([MS-OXCROPS] section 2.2.5.8), except that it moves the cursor using a custom bookmark, as specified in section 2.2.2.1.2, as a reference.");
            ModelHelper.CaptureRequirement(522, @"[In Processing RopCreateBookmark] When the server receives a RopCreateBookmark ROP request ([MS-OXCROPS] section 2.2.5.11), it MUST create a custom bookmark, as specified in section 2.2.2.1.2, that uniquely identifies a row in the table and can be subsequently used in the RopSeekRowBookmark ROP ([MS-OXCROPS] section 2.2.5.9).");
            return TableRopReturnValues.success;
        }

        /// <summary>
        /// This method is used to handle value which RopSeekRowBookmarkRequest returned
        /// </summary>
        /// <param name="isWantRowMovedCountout">Identify whether to return the actual seek row count</param>
        /// <param name="isEnoughRow">Identify whether there are enough row to seek</param>
        /// <param name="hasSoughtLess">Identify whether the actual seek row is less the one requested</param>
        /// <param name="isActualRowsCount">Identify whether the actual seek row is correct</param>
        /// <param name="rowNoLongerVisible">Identify whether the row is visible</param>
        [Rule(Action = "RopSeekRowBookmarkResponse(isWantRowMovedCountout,isEnoughRow,hasSoughtLess,isActualRowsCount,rowNoLongerVisible)")]
        public static void RopSeekRowBookmarkResponse(bool isWantRowMovedCountout, bool isEnoughRow, bool hasSoughtLess, bool isActualRowsCount, bool rowNoLongerVisible)
        {
            if (isWantRowMovedCountout)
            {
                if (isEnoughRow)
                {
                    {
                        Condition.IsTrue(isActualRowsCount && !hasSoughtLess);
                        ModelHelper.CaptureRequirement(178, @"[In RopSeekRowBookmark ROP Request Buffer] If this field [WantRowMovedCount] is set to ""TRUE"" (0x01), the server MUST return the actual number of rows moved.");
                        ModelHelper.CaptureRequirement(503, @"[In Processing RopSeekRowBookmark] The RopSeekRowBookmark ROP ([MS-OXCROPS] section 2.2.5.9) MUST move the cursor position according to its request fields.");
                        ModelHelper.CaptureRequirement(188, @"[In RopSeekRowBookmark ROP Response Buffer] Otherwise [if the number of rows moved is not less than the number of rows requested in the RowCount field.] it [HasSoughtLess] MUST be set to ""FALSE"" (0x00).");
                
                        // In the request, it uses adapter to specify a pre-defined bookmark and the number of 
                        // rows to move (forward or backwards), so, if the IsActualRowscount is true, MS-OXCTABL_R166 and MS-OXCTABL_R190
                        // can be verified.
                        ModelHelper.CaptureRequirement(166, "[In RopSeekRowBookmark ROP] The new location is specified by a custom bookmark, as specified in section 2.2.2.1.2, and the number of rows to move (forward or backwards) from that bookmark.");
                        ModelHelper.CaptureRequirement(190, @"[In RopSeekRowBookmark ROP Response Buffer] This field's [HasSoughtLess's] value MUST be valid if the WantRowMovedCount field (in the request) is set to ""TRUE"" (0x01).");
                    }
                }
                else
                {
                    Condition.IsTrue(isActualRowsCount && hasSoughtLess);
                    ModelHelper.CaptureRequirement(180, @"[In RopSeekRowBookmark ROP Request Buffer] The actual number of rows moved can differ from the requested number of rows if the beginning or end of the table is encountered before moving the requested number of rows.");
                    ModelHelper.CaptureRequirement(187, @"[In RopSeekRowBookmark ROP Response Buffer] This field [HasSoughtLess] MUST be set to ""TRUE"" (0x01) if the number of rows moved is less than the number of rows requested in the RowCount field.");
                }
            }

            // RowNoLongerVisible field MUST be set to FALSE (0x00) if the row to which the bookmark pointed 
            // has not been removed from the table
            if (validBookmark && !rowNoLongerVisible)
            {
                ModelHelper.CaptureRequirement(183, @"[In RopSeekRowBookmark ROP Response Buffer] Otherwise, [if the row to which the bookmark pointed has not been removed from the table] This field [RowNoLongerVisible] MUST be set to ""FALSE"" (0x00).");
            }
        }
        #endregion

        #region RopSeekRowFractional
        /// <summary>
        /// This method is used to seek row fractional for a table
        /// </summary>
        /// <param name="isZeroNumerator">Identify whether the Numerator is zero</param>
        /// <param name="isSmallDenominator">Identify whether the Numerator bigger than Denominator</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopSeekRowFractional(isZeroNumerator,isSmallDenominator)/result")]
        public static TableRopReturnValues RopSeekRowFractional(bool isZeroNumerator, bool isSmallDenominator)
        {
            // This ROP is only valid on Table objects.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                // The RopSeekRowFractional ROP MUST be supported for contents tables.
                ModelHelper.CaptureRequirement(855, @"[In Processing RopSeekRowFractional] The RopSeekRowFractional ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // The RopSeekRowFractional ROP MUST be supported for hierarchy tables.
                ModelHelper.CaptureRequirement(853, @"[In Processing RopSeekRowFractional] The RopSeekRowFractional ROP MUST be supported for hierarchy tables.");
            }
            else if (globleTableType == TableType.RULES_TABLE)
            {
                // The RopSeekRowFractional ROP MUST be supported for attachments tables.
                ModelHelper.CaptureRequirement(854, @"[In Processing RopSeekRowFractional] The RopSeekRowFractional ROP MUST be supported for rules tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(521, @"[In Processing RopSeekRowFractional] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not of type table.");

                ModelHelper.CaptureRequirement(200, @"[In RopSeekRowFractional ROP] This ROP is valid only on Table objects.");

                return TableRopReturnValues.ecNotSupported;
            }

            // Exchange 2010 SP1 and above implement the RopSeekRowFractional ROP.
            ModelHelper.CaptureRequirement(910, @"[In Appendix A: Product Behavior] Implementation does support the RopSeekRowFractional ROP. (<14> Section 2.2.2.11: Exchange 2010 SP1 and Exchange 2013, and Exchange 2016 do implement the RopSeekRowFractional ROP.)");
                
            return TableRopReturnValues.success;
        }

        /// <summary>
        /// This method is used to handle value which RopSeekRowFractionalRequest returned
        /// </summary>
        /// <param name="isZeroNumerator">Identify whether the Numerator is zero</param>
        /// <param name="isSmallDenominator">Identify whether the Numerator bigger than Denominator</param>
        /// <param name="cursorPosition">Cursor position</param>
        [Rule(Action = "RopSeekRowFractionalResponse(isZeroNumerator,isSmallDenominator,cursorPosition)")]
        public static void RopSeekRowFractionalResponse(bool isZeroNumerator, bool isSmallDenominator, CursorPosition cursorPosition)
        {
            // When numerator field is 0, the cursor position is begin
            if (isZeroNumerator)
            {
                Condition.IsTrue(cursorPosition == CursorPosition.BEGIN);
                ModelHelper.CaptureRequirement(514, @"[In Processing RopSeekRowFractional] If the Numerator field is set to 0, the cursor MUST move to the beginning of the table.");
            }

            // When numerator field is greater than denominator, the cursor position is end
            if (isSmallDenominator)
            {
                Condition.IsTrue(cursorPosition == CursorPosition.END);
                ModelHelper.CaptureRequirement(515, @"[In Processing RopSeekRowFractional] If the Numerator field value is greater than or equal to the Denominator field, the cursor MUST move to the end of the table.");
            }
        }
        #endregion

        #region RopCreateBookmark
        /// <summary>
        /// This method is used to create bookmark for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopCreateBookmark()/result")]
        public static TableRopReturnValues RopCreateBookmark()
        {
            // This ROP is valid only on contents table and hierarchy table.
            if (globleTableType == TableType.RULES_TABLE || globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                ModelHelper.CaptureRequirement(529, @"[In Processing RopCreateBookmark] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not a contents table or hierarchy table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else if (globleTableType == TableType.CONTENT_TABLE)
            {
                // The RopCreateBookmark ROP MUST be supported for contents tables.
                ModelHelper.CaptureRequirement(526, @"[In Processing RopCreateBookmark] The RopCreateBookmark ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // The RopCreateBookmark ROP MUST be supported for hierarchy tables.
                ModelHelper.CaptureRequirement(857, @"[In Processing RopCreateBookmark] The RopCreateBookmark ROP MUST be supported for hierarchy tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(209, @"[In RopCreateBookmark ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            validBookmark = true;
            bookmarkCreated = true;
            return TableRopReturnValues.success;
        }
        #endregion

        #region RopQueryColumnsAll
        /// <summary>
        /// This method is used to query all columns for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopQueryColumnsAll()/result")]
        public static TableRopReturnValues RopQueryColumnsAll()
        {
            // This ROP is only valid on Table objects.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                // The RopQueryRows ROP MUST be supported for contents tables.
                ModelHelper.CaptureRequirement(876, @"[In Processing RopQueryColumnsAll] The RopQueryColumnsAll ROP MUST be supported on contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // The RopQueryRows ROP MUST be supported for hierarchy tables.
                ModelHelper.CaptureRequirement(875, @"[In Processing RopQueryColumnsAll] The RopQueryColumnsAll ROP MUST be supported on hierarchy tables.");
            }
            else if (globleTableType == TableType.RULES_TABLE)
            {
                // The RopQueryRows ROP MUST be supported for attachments tables.
                ModelHelper.CaptureRequirement(877, @"[In Processing RopQueryColumnsAll] The RopQueryColumnsAll ROP MUST be supported on rules tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(533, @"[In Processing RopQueryColumnsAll] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80 , if the object on which this ROP was sent is not of type table.");
                ModelHelper.CaptureRequirement(218, @"[In RopQueryColumnsAll ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            return TableRopReturnValues.success;
        }
        #endregion

        #region RopFindRow
        /// <summary>
        /// This method is used to find row for a table
        /// </summary>
        /// <param name="findRowFlags">Identify the flags for this operation</param>
        /// <param name="origin">Identify the bookmark</param>
        /// <param name="isValidRestriction">Identify whether the restriction is valid</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopFindRow(findRowFlags,origin,isValidRestriction)/result")]
        public static TableRopReturnValues RopFindRow(FindRowFlags findRowFlags, BookmarkType origin, bool isValidRestriction)
        {
            // This ROP is valid only on contents table, rules table and hierarchy table.
            if (globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                ModelHelper.CaptureRequirement(544, @"[In Processing RopFindRow] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not a contents table, hierarchy table, or rules table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else if (globleTableType == TableType.CONTENT_TABLE)
            {
                // The RopFindRow ROP MUST be supported on contents tables.
                ModelHelper.CaptureRequirement(859, @"[In Processing RopFindRow] The RopFindRow ROP MUST be supported on contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // The RopFindRow ROP MUST be supported on hierarchy tables.
                ModelHelper.CaptureRequirement(858, @"[In Processing RopFindRow] The RopFindRow ROP MUST be supported on hierarchy tables.");
            }
            else if (globleTableType == TableType.RULES_TABLE)
            {
                // The RopFindRow ROP MUST be supported on rules tables.
                ModelHelper.CaptureRequirement(874, @"[In Processing RopFindRow] The RopFindRow ROP MUST be supported on rules tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(682, @"[In RopFindRow ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            // Before a success SetColumnsRop, FindRowRop will fail in Exchange2010 and above
            if (!setColumnsFlags.ContainsValue(true))
            {
                if (requirementContainer[827])
                {
                    ModelHelper.CaptureRequirement(827, @"[In Appendix A: Product Behavior] If a RopFindRow ROP is sent before a successful RopSetColumns ROP, then the implementation fails the ROP with ""ecNullObject"". (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    return TableRopReturnValues.ecNullObject;
                }
                else
                {
                    return TableRopReturnValues.unexpected;
                }
            }

            if (!validBookmark && bookmarkCreated)
            {
                // After resetTable Rop, attempts to use the bookmark will fail with ecInvalidBookmark except Exchange 2007.
                if (resetTableDone)
                {
                    Condition.IsTrue((requirementContainer[902] && !requirementContainer[903]) || (!requirementContainer[902] && requirementContainer[903]));
                    if (requirementContainer[902])
                    {
                        ModelHelper.CaptureRequirement(902, @"[In Appendix A: Product Behavior] Implementation does set the ReturnValue field to ""ecInvalidBookmark"" [for RopFindRow], if the client requested that the find be performed from a custom bookmark, as specified in section 2.2.2.1.2, but the bookmark has become invalid because of a RopResetTable ([MS-OXCROPS] section 2.2.5.15) ROP request. (Exchange Server 2010 and above follow this behavior.)");
                        return TableRopReturnValues.ecInvalidBookmark;
                    }

                    if (requirementContainer[903])
                    {
                        ModelHelper.CaptureRequirement(903, @"[In Appendix A: Product Behavior] Implementation returns ecSuccess [for RopFindRow], if the client requested that the find be performed from a custom bookmark, as specified in section 2.2.2.1.2, but the bookmark has become invalid because of a RopResetTable ([MS-OXCROPS] section 2.2.5.15) ROP request. (<33> Section 3.2.5.14: Exchange 2007 returns ecSuccess.)");
                    }
                }

                // After sortTable Rop, attempts to use the bookmark will fail with ecInvalidBookmark except Exchange 2007.
                if (sortTableFlags.ContainsValue(true))
                {
                    Condition.IsTrue((requirementContainer[898] && !requirementContainer[899]) || (!requirementContainer[898] && requirementContainer[899]));
                    if (requirementContainer[898])
                    {
                        ModelHelper.CaptureRequirement(898, @"[In Appendix A: Product Behavior] Implementation does set the ReturnValue field to ""ecInvalidBookmark"" [for RopFindRow], if the client requested that the find be performed from a custom bookmark, as specified in section 2.2.2.1.2, but the bookmark has become invalid because of a RopSortTable ([MS-OXCROPS] section 2.2.5.2) ROP request. (Exchange Server 2010 and above follow this behavior.)");
                        return TableRopReturnValues.ecInvalidBookmark;
                    }

                    if (requirementContainer[899])
                    {
                        ModelHelper.CaptureRequirement(899, @"[In Appendix A: Product Behavior] Implementation returns ecSuccess [for RopFindRow], if the client requested that the find be performed from a custom bookmark, as specified in section 2.2.2.1.2, but the bookmark has become invalid because of a RopSortTable ([MS-OXCROPS] section 2.2.5.2) ROP request. (<33> Section 3.2.5.14: Exchange 2007 returns ecSuccess.)");
                    }
                }

                // After restrict Table Rop, attempts to use the bookmark will fail with ecInvalidBookmark in Exchange 2010 and above.
                if (restrictFlags.ContainsValue(true))
                {
                    if (requirementContainer[900])
                    {
                        ModelHelper.CaptureRequirement(900, @"[In Appendix A: Product Behavior] Implementation does set the ReturnValue field to ""ecInvalidBookmark"" [for RopFindRow], if the client requested that the find be performed from a custom bookmark, as specified in section 2.2.2.1.2, but the bookmark has become invalid because of a RopRestrict ([MS-OXCROPS] section 2.2.5.3) ROP request. (Exchange Server 2010 and above follow this behavior.)");
                        return TableRopReturnValues.ecInvalidBookmark;
                    }
                }
            }

            return TableRopReturnValues.success;
        }

        /// <summary>
        /// This method is used to handle value which RopFindRowRequest returned
        /// </summary>
        /// <param name="findRowFlags">Identify the flags for this operation</param>
        /// <param name="isValidRestriction">Identify whether the restriction is valid to the table, that means there will be at least one data will be found</param>
        /// <param name="hasRowData">Identify whether there are rows meet the restriction, this is returned by server response</param>
        /// <param name="isCorrectRowData">Identify whether row returned meets the restriction</param>
        /// <param name="rowNoLongerVisible">Identify whether the row which is pointed by bookmark is visible</param>
        /// <param name="isLastSetColumnData">Identify whether the column data returned is the one set in the SetColumn method</param>
        [Rule(Action = "RopFindRowResponse(findRowFlags,isValidRestriction,hasRowData,isCorrectRowData,rowNoLongerVisible,isLastSetColumnData)")]
        public static void RopFindRowResponse(FindRowFlags findRowFlags, bool isValidRestriction, bool hasRowData, bool isCorrectRowData, bool rowNoLongerVisible, bool isLastSetColumnData)
        {
            if (requirementContainer[608])
            {
                if (globleTableType == TableType.HIERARCHY_TABLE)
                {
                    Condition.IsTrue(!rowNoLongerVisible);
                    ModelHelper.CaptureRequirement(608, @"[In Appendix A: Product Behavior] Implementation does always set this value [RowNoLongerVisible] to ""FALSE"" (0x00) for hierarchy tables in RopFindRow. (<15> Section 2.2.2.14.2: Exchange 2007 and Exchange 2010 always set this value [RowNoLongerVisible] to ""FALSE"" (0x00) for hierarchy tables in RopFindRow.)");
                }
            }

            // RopFindRow sets the cursor position to the first row that matches the criteria specified in the ROP (starting the search from the current cursor position) and returns the found row
            if (isValidRestriction && hasRowData)
            {
                Condition.IsTrue(isCorrectRowData);
                ModelHelper.CaptureRequirement(534, @"[In Processing RopFindRow]The RopFindRow ROP ([MS-OXCROPS] section 2.2.5.13) sets the cursor position to the first row that matches the search criteria specified in the ROP (starting the search from the current cursor position) and returns the found row when there is enough space in the output buffer.");
                ModelHelper.CaptureRequirement(243, @"[In RopFindRow ROP Response Buffer] If a row that meets the specified search criteria was found and the row data is included in the response, this field [HasRowData] MUST be set to ""TRUE"" (0x01).");
            }

            // The columns sent for the row found MUST be the columns that are specified on RopSetColumns
            if (setColumnsFlags.ContainsValue(true))
            {
                if (!resetTableDone)
                {
                    Condition.IsTrue(isLastSetColumnData);
                    ModelHelper.CaptureRequirement(540, @"[In Processing RopFindRow] The columns sent for the row found MUST be the columns that are specified on the RopSetColumns ROP.");
                    ModelHelper.CaptureRequirement(539, @"[In Processing RopFindRow]A RopSetColumns ROP request ([MS-OXCROPS] section 2.2.5.1) MUST be sent on the table before sending a RopFindRow ROP request.");
                    ModelHelper.CaptureRequirement(813, @"[In RopSetColumns ROP Request Buffer] PropertyTags (variable): An array of PropertyTag structures, as described in [MS-OXCDATA] section 2.9, that identify the set and order of property values to be returned by the server in the ROP response buffer of the RopFindRow ([MS-OXCROPS] section 2.2.5.13)ROP, as specified in section 2.2.2.14.");
                    ModelHelper.CaptureRequirement(824, @"[In Processing RopSetColumns] The columns that are set by the RopSetColumns ROP MUST be the ones sent in the responses to subsequent RopFindRow ([MS-OXCROPS] section 2.2.5.13) ROP executed on that table.");
                }
            }

            if (findRowFlags == FindRowFlags.Forwards)
            {
                // If the value of FindRowFlags is set to 0x00 in the request and the correct row can be found, this requirement can be verified.
                Condition.IsTrue(isCorrectRowData);
                ModelHelper.CaptureRequirement(225, @"[In RopFindRow ROP Request Buffer] The value of FindRowFlags is 0x00 means perform the find forwards.");
            }
            else if (findRowFlags == FindRowFlags.Backwards)
            {
                // If the value of findRowFlags is set to 0x01 in the request and the correct row can be found, this requirement can be verified.
                Condition.IsTrue(isCorrectRowData);
                ModelHelper.CaptureRequirement(226, @"[In RopFindRow ROP Request Buffer] The value of FindRowFlags is 0x01 means perform the find backwards.");
            }

            if (validBookmark)
            {
                Condition.IsTrue(!rowNoLongerVisible);
                ModelHelper.CaptureRequirement(239, @"[In RopFindRow ROP Response Buffer] Otherwise [if the row to which the bookmark pointed, has not been removed from the table.], this field [RowNoLongerVisible] MUST be set to ""FALSE"" (0x00).");
            }
        }
        #endregion

        #region RopFreeBookmark
        /// <summary>
        /// This method is used to free bookmark for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopFreeBookmark()/result")]
        public static TableRopReturnValues RopFreeBookmark()
        {
            // This ROP is valid only on contents table and hierarchy table.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                // RopFreeBookmark SHOULD be supported by contents tables for Microsoft Exchange Server 2010 and above.
                if (requirementContainer[549])
                {
                    ModelHelper.CaptureRequirement(549, @"[In Appendix A: Product Behavior] The RopFreeBookmark ROP is supported by contents tables on the implementation. (Microsoft Exchange Server 2010 and above follow this behavior.)");
                }
                else
                {
                    return TableRopReturnValues.unexpected;
                }
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE)
            {
                // RopFreeBookmark SHOULD be supported by hierarchy tables.
                ModelHelper.CaptureRequirement(547, @"[In Processing RopFreeBookmark] The RopFreeBookmark ROP MUST be supported by hierarchy tables.");
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project                  
                ModelHelper.CaptureRequirement(249, @"[In RopFreeBookmark ROP] This ROP is valid only on Table objects.");
                ModelHelper.CaptureRequirement(551, @"[In Processing RopFreeBookmark] The error code ecNotSupported will be returned with value 0x80040102 %x02.01.04.80, if the object on which this ROP was sent is not a hierarchy table.");
                return TableRopReturnValues.ecNotSupported;
            }

            // Before a success SetColumnsRop, RopFreeBookmark will fail in Exchange2010 and above
            if (!setColumnsFlags.ContainsValue(true) && !resetTableDone)
            {
                if (requirementContainer[8274])
                {
                    ModelHelper.CaptureRequirement(8274, @"[In Appendix A: Product Behavior] If a RopFreeBookmark ROP is sent before a successful RopSetColumns ROP, then the implementation fails the ROP with ""ecNullObject"". (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    return TableRopReturnValues.ecNullObject;
                }
                else
                {
                    return TableRopReturnValues.unexpected;
                }
            }

            if (!validBookmark)
            {
                // After resetTable Rop, attempts to use the bookmark will fail with ecNullObject.
                if (resetTableDone)
                {
                    ModelHelper.CaptureRequirement(863, @"[In Processing RopCreateBookmark] If the client does not send a RopFreeBookmark ROP request ([MS-OXCROPS] section 2.2.5.14), the server MUST release all bookmarks related to a table when a RopResetTable ([MS-OXCROPS] section 2.2.5.15) ROP request is sent.");
                    ModelHelper.CaptureRequirement(562, @"[In Processing RopResetTable] After a RopResetTable ROP executes, all previously existing bookmarks on the table are invalid.");
                    return TableRopReturnValues.ecNullObject;
                }

                // After sortTable Rop, attempts to use the bookmark will fail with ecInvalidBookmark.
                if (sortTableFlags.ContainsValue(true))
                {
                    ModelHelper.CaptureRequirement(864, @"[In Processing RopCreateBookmark] If the client does not send a RopFreeBookmark ROP request ([MS-OXCROPS] section 2.2.5.14), the server MUST release all bookmarks related to a table when a RopSortTable ([MS-OXCROPS] section 2.2.5.2) ROP request is sent.");
                }

                // After restrict Table Rop, attempts to use the bookmark will fail with ecInvalidBookmark.
                if (restrictFlags.ContainsValue(true))
                {
                    ModelHelper.CaptureRequirement(865, @"[In Processing RopCreateBookmark] If the client does not send a RopFreeBookmark ROP request ([MS-OXCROPS] section 2.2.5.14), the server MUST release all bookmarks related to a table when a RopRestrict ([MS-OXCROPS] section 2.2.5.3) ROP request is sent.");
                }

                ModelHelper.CaptureRequirement(545, @"[In Processing RopFreeBookmark] The RopFreeBookmark ROP ([MS-OXCROPS] section 2.2.5.14) MUST release any resources on the server used to keep track of the bookmark (created using a RopCreateBookmark ROP ([MS-OXCROPS] section 2.2.5.11)).");
                return TableRopReturnValues.ecInvalidBookmark;
            }
            else if (validBookmark)
            {
                validBookmark = false;
            }

            // Set the free bookmark flag
            freeBookmarkDone = true;

            return TableRopReturnValues.success;
        }
        #endregion

        #region RopResetTable
        /// <summary>
        /// This method is used to reset a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopResetTable()/result")]
        public static TableRopReturnValues RopResetTable()
        {
            // This ROP is valid only on contents table.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                ModelHelper.CaptureRequirement(564, @"[In Processing RopResetTable] The RopResetTable ROP MUST be supported on contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE || globleTableType == TableType.RULES_TABLE || globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                // The error code ecNotSupported will be returned with value 0x80040102(%x02.01.04.80), if the object on which this ROP was sent is not a contents table.
                ModelHelper.CaptureRequirement(566, @"[In Processing RopResetTable] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not a contents table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(259, @"[In RopResetTable ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            // If a bookmark has been created, invalid it.
            if (validBookmark)
            {
                validBookmark = false;
            }

            // Set the reset table flag
            resetTableDone = true;

            // Clear the set columns flags, this is used to make sure that when the reset ROP is done, the SetColumn must be
            // sent before the RopFindRow, RopQueryRows, or RopExpandRow.
            setColumnsFlags.Clear();
            return TableRopReturnValues.success;
        }

        /// <summary>
        /// This method is used to handle value which RopResetTableRequest returned
        /// </summary>
        /// <param name="ropType">Identify ROP type</param>
        /// <param name="isPreviousDataReset">Identify whether the previous data is reset</param>
        [Rule(Action = "RopResetTableResponse(ropType,isPreviousDataReset)")]
        public static void RopResetTableResponse(TableRopType ropType, bool isPreviousDataReset)
        {
            if (isPreviousDataReset)
            {
                switch (ropType)
                {
                    // The previous restriction is reset
                    case TableRopType.RESTRICT:
                        {
                            ModelHelper.CaptureRequirement(555, @"[In Processing RopResetTable] The RopResetTable ROP MUST remove the restriction previously applied to the table using RopRestrict ROP ([MS-OXCROPS] section 2.2.5.3) (if any).");
                            ModelHelper.CaptureRequirement(556, @"[In Processing RopResetTable] The table MUST afterwards appear as if the RopRestrict ROP had never been sent on it; that is, as if it had no restriction (all rows MUST be present).");
                            break;
                        }

                    // The cursor position is reset
                    case TableRopType.MOVECURSOR:
                        {
                            ModelHelper.CaptureRequirement(561, @"[In Processing RopResetTable] The RopResetTable ROP MUST move the cursor to the beginning of the table.");
                            break;
                        }

                    // The created bookmark is invalidated
                    case TableRopType.CREATEBOOKMARK:
                        {
                            break;
                        }

                    default:
                        break;
                }
            }
        }
        #endregion

        #region RopExpandRow
        /// <summary>
        /// This method is used to expand row for a table
        /// </summary>
        /// <param name="nonexistCategoryId">Identify whether the CategoryId row exists</param>
        /// <param name="noncollapsedCategoryId">Identify whether the CategoryId row is not collapsed</param>
        /// <param name="isMaxRowCountInvalid">Identify whether the MaxRowCount is invalid</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopExpandRow(nonexistCategoryId,noncollapsedCategoryId, isMaxRowCountInvalid)/result")]
        public static TableRopReturnValues RopExpandRow(bool nonexistCategoryId, bool noncollapsedCategoryId, bool isMaxRowCountInvalid)
        {
            if (isMaxRowCountInvalid)
            {
                if (requirementContainer[748])
                {
                    ModelHelper.CaptureRequirement(748, @"[In Appendix A: Product Behavior] Implementation does not support a value greater than 0 for the MaxRowCount field. (<16> Section 2.2.2.17.1: Exchange 2013 and Exchange 2016 do not support a value greater than 0 for the MaxRowCount field.)");
                    return TableRopReturnValues.ecNotSupported;
                }
            }

            // This ROP is valid only on contents table.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                ModelHelper.CaptureRequirement(568, @"[In Processing RopExpandRow] The RopExpandRow ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE || globleTableType == TableType.RULES_TABLE || globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                // The error code ecNotSupported will be returned with value 0x80040102(%x02.01.04.80), if the object on which this ROP was sent is not a contents table.
                ModelHelper.CaptureRequirement(572, @"[In Processing RopExpandRow] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not a contents table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(274, @"[In RopExpandRow ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            // The error ecNotFound will be returned with value 0x8004010F(%x0F.01.04.80),
            // means the row specified by the CategoryId field was not found
            if (nonexistCategoryId)
            {
                ModelHelper.CaptureRequirement(570, @"[In Processing RopExpandRow] The error ecNotFound will be returned with value 0x8004010F,%x0F.01.04.80, if the row specified by the CategoryId field was not found.");
                return TableRopReturnValues.ecNotFound;
            }

            // The error code ecNotCollapsed will be returned with value 0x000004F8(%xF8.04.00.00), 
            // means the row specified by the CategoryId field was not collapsed
            if (noncollapsedCategoryId)
            {
                ModelHelper.CaptureRequirement(571, @"[In Processing RopExpandRow] The error code ecNotCollapsed will be returned with value 0x000004F8,%xF8.04.00.00 means the row specified by the CategoryId field was not collapsed.");
                return TableRopReturnValues.ecNotCollapsed;
            }

            // Before a success SetColumnsRop, RopExpandRow will fail in Exchange2010 and above
            if (!setColumnsFlags.ContainsValue(true))
            {
                if (requirementContainer[8273])
                {
                    ModelHelper.CaptureRequirement(8273, @"[In Appendix A: Product Behavior] If a RopExpandRow ROP is sent before a successful RopSetColumns ROP, then the implementation fails the ROP with ""ecNullObject"". (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    return TableRopReturnValues.ecNullObject;
                }
                else
                {
                    return TableRopReturnValues.unexpected;
                }
            }

            ModelHelper.CaptureRequirement(567, @"[In Processing RopExpandRow] The RopExpandRow ROP ([MS-OXCROPS] section 2.2.5.16) sets a category row to expanded state.");
            return TableRopReturnValues.success;
        }

        /// <summary>
        /// This method is used to handle value which RopExpandRowRequest returned
        /// </summary>
        /// <param name="isLastSetColumnData">Identify whether the column data returned is the one set in the SetColumn method</param>
        [Rule(Action = "RopExpandRowResponse(isLastSetColumnData)")]
        public static void RopExpandRowResponse(bool isLastSetColumnData)
        {
            // The columns sent for the row found MUST be the columns that are specified on RopSetColumns
            if (setColumnsFlags.ContainsValue(true))
            {
                if (!resetTableDone)
                {
                    // The ExpandRow ROP will not return any PropertyTag structure under Exchange server 2013 
                    // since it does not support the MaxRowCount field in the request to be set greater than 0.
                    if (!requirementContainer[748])
                    {
                        Condition.IsTrue(isLastSetColumnData);
                        ModelHelper.CaptureRequirement(814, @"[In RopSetColumns ROP Request Buffer] PropertyTags (variable): An array of PropertyTag structures, as described in [MS-OXCDATA] section 2.9, that identify the set and order of property values to be returned by the server in the ROP response buffer of the RopExpandRow ([MS-OXCROPS] section 2.2.5.16) ROP, as specified in section 2.2.2.17.");
                        ModelHelper.CaptureRequirement(825, @"[In Processing RopSetColumns] The columns that are set by the RopSetColumns ROP MUST be the ones sent in the responses to subsequent  RopExpandRow ([MS-OXCROPS] section 2.2.5.16) ROP executed on that table.");
                    }
                }
            }
        }
        #endregion

        #region RopCollapseRow
        /// <summary>
        /// This method is used to Collapse row for a table
        /// </summary>
        /// <param name="nonexistCategoryId">Identify whether the CategoryId row exists</param>
        /// <param name="nonexpandedCategoryId">Identify whether the CategoryId row is not expanded</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopCollapseRow(nonexistCategoryId,nonexpandedCategoryId)/result")]
        public static TableRopReturnValues RopCollapseRow(bool nonexistCategoryId, bool nonexpandedCategoryId)
        {
            // This ROP is valid only on contents table.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                ModelHelper.CaptureRequirement(574, @"[In Processing RopCollapseRow] The RopCollapseRow ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE || globleTableType == TableType.RULES_TABLE || globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                ModelHelper.CaptureRequirement(578, @"[In Processing RopCollapseRow] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not a contents table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(284, @"[In RopCollapseRow ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            // The error code ecNotFound will be returned with value 0x8004010F(%x0F.01.04.80), means 
            // the row specified by the CategoryId field was not found
            if (nonexistCategoryId)
            {
                ModelHelper.CaptureRequirement(576, @"[In Processing RopCollapseRow] The error code ecNotFound will be returned with value 0x8004010F,%x0F.01.04.80, if the row specified by the CategoryId field was not found.");
                return TableRopReturnValues.ecNotFound;
            }

            // The error code ecNotExpanded will be returned with value 0x000004F7(%xF7.04.00.00),
            // means the row specified by the CategoryId field was not expanded
            if (nonexpandedCategoryId)
            {
                ModelHelper.CaptureRequirement(577, @"[In Processing RopCollapseRow] The error code ecNotExpanded will be returned with value 0x000004F7,%xF7.04.00.00 means the row specified by the CategoryId field was not expanded.");
                ModelHelper.CaptureRequirement(573, @"[In Processing RopCollapseRow] The RopCollapseRow ROP ([MS-OXCROPS] section 2.2.5.17) MUST set a category row to collapsed state.");
                return TableRopReturnValues.ecNotExpanded;
            }

            // Before a success SetColumnsRop, RopCollapseRow will fail in Exchange2010 and above
            if (!setColumnsFlags.ContainsValue(true))
            {
                if (requirementContainer[8271])
                {
                    ModelHelper.CaptureRequirement(8271, @"[In Appendix A: Product Behavior] If a RopCollapseRow ROP is sent before a successful RopSetColumns ROP, then the implementation fails the ROP with ""ecNullObject"". (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    return TableRopReturnValues.ecNullObject;
                }
                else
                {
                    return TableRopReturnValues.unexpected;
                }
            }

            return TableRopReturnValues.success;
        }
        #endregion

        #region RopGetCollapseState
        /// <summary>
        /// This method is used to get collapse state for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopGetCollapseState()/result")]
        public static TableRopReturnValues RopGetCollapseState()
        {
            // This ROP is valid only on contents table.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                ModelHelper.CaptureRequirement(582, @"[In Processing RopGetCollapseState]The RopGetCollapseState ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE || globleTableType == TableType.RULES_TABLE || globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                // The error code ecNotSupported will be returned if the object on which this ROP was sent is not a contents table
                ModelHelper.CaptureRequirement(584, @"[In Processing RopGetCollapseState] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80,if the object on which this ROP was sent is not a contents table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(290, @"[In RopGetCollapseState ROP] This ROP is valid only on Table objects.");
                return TableRopReturnValues.ecNotSupported;
            }

            // Before a success SetColumnsRop, RopGetCollapseState will fail in Exchange2010 and above
            if (!setColumnsFlags.ContainsValue(true))
            {
                if (requirementContainer[8275])
                {
                    ModelHelper.CaptureRequirement(8275, @"[In Appendix A: Product Behavior] If a RopGetCollapseState ROP is sent before a successful RopSetColumns ROP, then the implementation fails the ROP with ""ecNullObject"". (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    return TableRopReturnValues.ecNullObject;
                }
                else
                {
                    return TableRopReturnValues.unexpected;
                }
            }

            ModelHelper.CaptureRequirement(579, @"[In Processing RopGetCollapseState] The RopGetCollapseState ROP ([MS-OXCROPS] section 2.2.5.18) MUST send the collapsed state of the whole table in the CollapseState field of the ROP response.");
            return TableRopReturnValues.success;
        }

        /// <summary>
        /// This method is used to handle value which RopGetCollapseStateRequest returned
        /// </summary>
        /// <param name="isSetCollapseStateCalled">Identify whether the SetCollapseState Rop is called</param>
        /// <param name="isCollapseStateRebuildCorrectly">Identify whether the collapse state is rebuilt successfully</param>
        [Rule(Action = "RopGetCollapseStateResponse(isSetCollapseStateCalled,isCollapseStateRebuildCorrectly)")]
        public static void RopGetCollapseStateResponse(bool isSetCollapseStateCalled, bool isCollapseStateRebuildCorrectly)
        {
            if (isSetCollapseStateCalled)
            {
                Condition.IsTrue(isCollapseStateRebuildCorrectly);
                ModelHelper.CaptureRequirement(304, @"[In RopSetCollapseState ROP] The RopSetCollapseState ROP ([MS-OXCROPS] section 2.2.5.19) rebuilds a table's collapsed state, which is specified by the data returned from a RopGetCollapseState ROP ([MS-OXCROPS] section 2.2.5.18).");
            }
        }
        #endregion

        #region RopSetCollapseState
        /// <summary>
        /// This method is used to set collapse state for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopSetCollapseState()/result")]
        public static TableRopReturnValues RopSetCollapseState()
        {
            // This ROP is valid only on contents table.
            if (globleTableType == TableType.CONTENT_TABLE)
            {
                ModelHelper.CaptureRequirement(588, @"[In Processing RopSetCollapseState] The RopSetCollapseState ROP MUST be supported for contents tables.");
            }
            else if (globleTableType == TableType.HIERARCHY_TABLE || globleTableType == TableType.RULES_TABLE || globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
            {
                // The error code ecNotSupported will be returned if the object on which this ROP was sent is not a contents table
                ModelHelper.CaptureRequirement(590, @"[In Processing RopSetCollapseState] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80,if the object on which this ROP was sent is not a contents table.");
                return TableRopReturnValues.ecNotSupported;
            }
            else
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                ModelHelper.CaptureRequirement(306, "[In RopSetCollapseState ROP] This ROP is valid only on Table objects");
                return TableRopReturnValues.ecNotSupported;
            }

            // Before a success SetColumnsRop, RopSetCollapseState will fail in Exchange2010 and above
            if (!setColumnsFlags.ContainsValue(true))
            {
                if (requirementContainer[8277])
                {
                    ModelHelper.CaptureRequirement(8277, @"[In Appendix A: Product Behavior] If a RopSetCollapseState ROP is sent before a successful RopSetColumns ROP, then the implementation fails the ROP with ""ecNullObject"". (Microsoft Exchange Server 2010 and above follow this behavior.)");
                    return TableRopReturnValues.ecNullObject;
                }
                else
                {
                    return TableRopReturnValues.unexpected;
                }
            }

            // The RopGetCollapseState must be called before RopSetCollapseState, and if RopSetCollapseState is successful, then a book mark is returned, so the following requirement can be verified.
            ModelHelper.CaptureRequirement(585, @"[In Processing RopSetCollapseState] The RopSetCollapseState ROP ([MS-OXCROPS] section 2.2.5.19) MUST modify the collapsed state of the table to match the collapsed state being sent.");
            ModelHelper.CaptureRequirement(289, @"[In RopGetCollapseState ROP] The data returned is in the form of an opaque BLOB that can be passed to a RopSetCollapseState ROP request ([MS-OXCROPS] section 2.2.5.19).");
            return TableRopReturnValues.success;
        }
        #endregion
        #endregion

        #region RopAbort
        /// <summary>
        /// This method is used to Abort ROPs for a table
        /// </summary>
        /// <param name="tableStatus">Identify the table status</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopAbort(out tableStatus)/result")]
        public static TableRopReturnValues RopAbort(out TableStatus tableStatus)
        {
            tableStatus = TableStatus.TblstatComplete;

            if (requirementContainer[791])
            {
                // The ROP is valid only on contents table and hierarchy table
                if (globleTableType == TableType.RULES_TABLE || globleTableType == TableType.ATTACHMENTS_TABLE || globleTableType == TableType.PERMISSIONS_TABLE)
                {
                    ModelHelper.CaptureRequirement(486, @"[In Processing RopAbort] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not a contents table or a hierarchy table.");
                    return TableRopReturnValues.ecNotSupported;
                }
                else if (globleTableType == TableType.CONTENT_TABLE)
                {
                    ModelHelper.CaptureRequirement(483, @"[In Processing RopAbort] The RopAbort ROP MUST be supported for contents tables.");
                }
                else if (globleTableType == TableType.HIERARCHY_TABLE)
                {
                    ModelHelper.CaptureRequirement(839, @"[In Processing RopAbort] The RopAbort ROP MUST be supported for hierarchy tables.");
                }
                else
                {
                    // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                    // is referred as a folderId, for details, see the table initial method in adapter project
                    ModelHelper.CaptureRequirement(124, @"[In RopAbort ROP] This ROP is valid only on Table objects.");
                    return TableRopReturnValues.ecNotSupported;
                }

                if (tableStatus == TableStatus.TblstatComplete)
                {
                    return TableRopReturnValues.ecUnableToAbort;
                }
            }
            else
            {
                return TableRopReturnValues.NotImplemented;
            }

            return TableRopReturnValues.success;
        }
        #endregion

        #region RopGetStatus
        /// <summary>
        /// This method is used to get the status for a table
        /// </summary>
        /// <param name="tableStatus">Identify the table status</param>
        /// <returns>Table ROP return value</returns>
        [Rule(Action = "RopGetStatus(out tableStatus)/result")]
        public static TableRopReturnValues RopGetStatus(out TableStatus tableStatus)
        {
            tableStatus = TableStatus.TblstatComplete;

            if (requirementContainer[792])
            {
                // If the table type is invalid, ecNotSupported will be returned, in this test suite, invalid table
                // is referred as a folderId, for details, see the table initial method in adapter project
                if (globleTableType == TableType.INVALID_TABLE)
                {
                    ModelHelper.CaptureRequirement(490, @"[In Processing RopGetStatus] The error code ecNotSupported will be returned with value 0x80040102,%x02.01.04.80, if the object on which this ROP was sent is not of type table.");
                    ModelHelper.CaptureRequirement(129, @"[In RopGetStatus ROP] This ROP is valid only on Table objects.");
                    return TableRopReturnValues.ecNotSupported;
                }

                // The ROP is valid only on contents table, hierarchy table and attachments tables
                if (globleTableType == TableType.CONTENT_TABLE)
                {
                    ModelHelper.CaptureRequirement(842, @"[In Processing RopGetStatus] The RopGetStatus ROP MUST be supported for contents tables.");
                }
                else if (globleTableType == TableType.HIERARCHY_TABLE)
                {
                    ModelHelper.CaptureRequirement(840, @"[In Processing RopGetStatus] The RopGetStatus ROP MUST be supported for hierarchy tables.");
                }
                else if (globleTableType == TableType.RULES_TABLE)
                {
                    ModelHelper.CaptureRequirement(841, @"[In Processing RopGetStatus] The RopGetStatus ROP MUST be supported for rules tables.");
                }
            }
            else
            {
                return TableRopReturnValues.NotImplemented;
            }

            return TableRopReturnValues.success;
        }
        #endregion

        #region RopRelease
        /// <summary>
        /// This method is used to release a table
        /// </summary>
        [Rule(Action = "RopRelease()")]
        public static void RopRelease()
        {
            ropReleaseDone = true;
        }
        #endregion

        #region Helpers

        /// <summary>
        /// Parse the query row response after set columns
        /// </summary>
        /// <param name="isLatestRopData">Identify whether the queried data equals to the last ROP data</param>
        /// <param name="isLastSuccessRopData">Identify whether the queried row data equals to the last successful ROP data</param>  
        private static void ParseQueryRowsAfterSetColumns(bool isLatestRopData, bool isLastSuccessRopData)
        {
            // If SetColumn fails, the server can restore the previous column set.
            bool latestSetColumnFlag = false;
            setColumnsFlags.TryGetValue((uint)setColumnsFlags.Count, out latestSetColumnFlag);

            // If there is successful SetColumn response
            if (latestSetColumnFlag)
            {
                if (isLastSuccessRopData)
                {
                    ModelHelper.CaptureRequirement(472, @"[In Processing RopQueryRows] The columns sent in the response for each row MUST be the ones specified on the RopSetColumns ROP request.");
                    ModelHelper.CaptureRequirement(427, @"[In Processing RopSetColumns] The columns that are set by the RopSetColumns ROP MUST be the ones sent in the responses to subsequent RopQueryRows ([MS-OXCROPS] section 2.2.5.4) ROP executed on that table.");
                    ModelHelper.CaptureRequirement(426, @"[In Processing RopSetColumns] When a RopSetColumns ROP request ([MS-OXCROPS] section 2.2.5.1) is received, the server MUST remember the requested columns and apply them to the table when executing other ROPs that act on that table.");

                    // Here it can partially verify MS-OXCTABL_R116 by verifying the same columns
                    ModelHelper.CaptureRequirement(116, @"[In RopQueryRows ROP Response Buffer] Each row MUST have the same columns and ordering of columns as specified in the last RopSetColumns ROP request ([MS-OXCROPS] section 2.2.5.1).");
                }
            }

            // The server will invalid the table column until the successful SetColumn.
            bool lastSetColumnFlag = false;
            setColumnsFlags.TryGetValue((uint)setColumnsFlags.Count - 1, out lastSetColumnFlag);
            if (latestSetColumnFlag && !lastSetColumnFlag && isLatestRopData)
            {
                if (requirementContainer[610])
                {
                    // If a RopSetColumns ROP request fails, the server invalidate the table column set until a successful RopSetColumns ROP request is made.
                    ModelHelper.CaptureRequirement(610, @"[In Appendix A: Product Behavior] If a RopSetColumns ROP request fails, the implementation does invalidate the table column set until a successful RopSetColumns ROP request is made. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// Parse the query row response after sort table
        /// </summary>
        /// <param name="isLatestRopData">Identify whether the queried data equals to the last ROP data</param>
        /// <param name="isLastSuccessRopData">Identify whether the queried row data equals to the last successful ROP data</param>  
        private static void ParseQueryRowsAfterSortTable(bool isLatestRopData, bool isLastSuccessRopData)
        {
            // If there is successful Sort Table response
            if (sortTableFlags.ContainsValue(true))
            {
                if (isLastSuccessRopData)
                {
                    ModelHelper.CaptureRequirement(437, @"[In Processing RopSortTable] When a RopSortTable ROP request ([MS-OXCROPS] section 2.2.5.2) is received, the server MUST apply the sort order to the table, and subsequent requests sent that operate on the table MUST consider the new sort order.");

                    // Here it can partially verify MS-OXCTABL_R116 by verifying the same order
                    ModelHelper.CaptureRequirement(116, @"[In RopQueryRows ROP Response Buffer] Each row MUST have the same columns and ordering of columns as specified in the last RopSetColumns ROP request ([MS-OXCROPS] section 2.2.5.1).");

                    ModelHelper.CaptureRequirement(66, @"[In RopSortTable ROP Request Buffer] The first categories of the ExpandedCount field are initially expanded.");
                    ModelHelper.CaptureRequirement(67, @"[In RopSortTable ROP Request Buffer] If the value of the CategoryCount field is equal to the value of the ExpandedCount field, then all categories are expanded.");
                    ModelHelper.CaptureRequirement(71, @"[In RopSortTable ROP Request Buffer] When the value of the SortOrderCount field exceeds the value of the CategoryCount field, indicating that there are more sort keys than categories, categories are created from the SortOrder structures that appear first in the SortOrders array.");
                    ModelHelper.CaptureRequirement(72, @"[In RopSortTable ROP Request Buffer] The remaining SortOrder structures are used to sort the rows within the categories.");
                    ModelHelper.CaptureRequirement(76, @"[In RopSortTable ROP Request Buffer] If the Order member of a SortOrder structure is set to ""Ascending"", the table will be sorted in ascending order by the column specified in the PropertyType and PropertyId members.");
                    ModelHelper.CaptureRequirement(77, @"[In RopSortTable ROP Request Buffer] If the Order member of a SortOrder structure is set to ""Descending"", the table will be sorted in descending order by the column specified in the PropertyType and PropertyId members.");
                }

                if (isLastSuccessRopData && isLatestRopData)
                {
                    ModelHelper.CaptureRequirement(119, @"[In RopQueryRows ROP Response Buffer] The rows MUST be sorted and grouped according to the sort order specified in the last RopSortTable ROP request ([MS-OXCROPS] section 2.2.5.2).");
                    ModelHelper.CaptureRequirement(437, @"[In Processing RopSortTable] When a RopSortTable ROP request ([MS-OXCROPS] section 2.2.5.2) is received, the server MUST apply the sort order to the table, and subsequent requests sent that operate on the table MUST consider the new sort order.");
                }
            }

            // If SortTable fails, the server can restore the previous order data
            bool latestSortTableFlag = false;
            sortTableFlags.TryGetValue((uint)sortTableFlags.Count, out latestSortTableFlag);
            if (latestSortTableFlag && isLastSuccessRopData && isLatestRopData)
            {
                ModelHelper.CaptureRequirement(438, @"[In Processing RopSortTable] If a sort order is already specified, the new sort order returned with the ROP response MUST replace the old sort order.");
            }

            // The server will invalid the table order until the successful SortTable.
            bool lastSortTableFlag = false;
            sortTableFlags.TryGetValue((uint)sortTableFlags.Count - 1, out lastSortTableFlag);

            if (!lastSortTableFlag && latestSortTableFlag && isLatestRopData)
            {
                // In this case, the latest SortTable success and there was failed record for SortTable, also 
                // if the row data equals to the latest sort data, following requirements can be verified.
                if (requirementContainer[443])
                {
                    ModelHelper.CaptureRequirement(443, @"[In Appendix A: Product Behavior] If the RopSortTable ROP fails, the implementation does invalidate the table sort order until a successful RopSortTable ROP is made. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// Parse the query row response after restrict table
        /// </summary>
        /// <param name="isLatestRopData">Identify whether the queried data equals to the last ROP data</param>
        /// <param name="isLastSuccessRopData">Identify whether the queried row data equals to the last successful ROP data</param>  
        private static void ParseQueryRowsAfterRestrict(bool isLatestRopData, bool isLastSuccessRopData)
        {
            // If there is successful restrict table response
            if (restrictFlags.ContainsValue(true))
            {
                if (isLastSuccessRopData)
                {
                    ModelHelper.CaptureRequirement(451, @"[In Processing RopRestrict] When a RopRestrict ROP request ([MS-OXCROPS] section 2.2.5.3) is received, the server MUST apply the restriction to the table, and subsequent requests that operate on the table MUST consider the new restriction.");
                    ModelHelper.CaptureRequirement(452, @"[In Processing RopRestrict] If a restriction is applied to a table, the table MUST appear as if it only contains the rows that match the restriction.");
                    ModelHelper.CaptureRequirement(117, @"[In RopQueryRows ROP Response Buffer] The RowData field MUST NOT include rows that don't match the criteria specified in the last RopRestrict ROP request ([MS-OXCROPS] section 2.2.5.3).");
                }
            }

            // If SetColumn fails, the server can restore the previous restriction data
            bool latestRestrictFlag = false;
            restrictFlags.TryGetValue((uint)restrictFlags.Count, out latestRestrictFlag);

            // The server will invalid the restrict data until the successful restrict data.
            bool lastRestrictFlag = false;
            restrictFlags.TryGetValue((uint)restrictFlags.Count - 1, out lastRestrictFlag);
            if (!lastRestrictFlag && latestRestrictFlag && isLatestRopData)
            {
                if (requirementContainer[456])
                {
                    ModelHelper.CaptureRequirement(456, @"[In Appendix A: Product Behavior] If a RopRestrict ROP fails, the implementation does invalidate the table restriction until a successful RopRestrict ROP request is made. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                }
            }
        }

        #endregion
    }
}