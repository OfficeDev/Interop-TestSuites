namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    #region Protocol rops handle

    /// <summary>
    /// This delegate is used to handle response which RopQueryRowsRequest returned
    /// </summary>
    /// <param name="queryRowFlags">A flag for the query row operation</param>
    /// <param name="bForwardRead">Identify whether to read the data forward</param>
    /// <param name="bZeroRow">Identify whether zero row returned</param>
    /// <param name="ropType">Identify rop type</param>
    /// <param name="bIsCorrectRowCount">Identify whether the correct row count data was read</param>
    /// <param name="bCursorPositionChanged">Identify whether the cursor position was changed</param>
    /// <param name="bIsLatestRopData">Identify whether the queried data equals to the last rop data</param>
    /// <param name="bIsLastSuccessRopData">Identify whether the queried row data equals to the last successful rop data</param>
    /// <param name="queryRowOrigin">Identify the cursor position returned in queryRow response</param>
    /// <param name="isRequestCountTooLarger">Identify whether the row count in the request is larger than the actual row count</param>
    public delegate void RopQueryRowsHandler(QueryRowsFlags queryRowFlags, bool bForwardRead, bool bZeroRow, TableRopType ropType, bool bIsCorrectRowCount, bool bCursorPositionChanged, bool bIsLatestRopData, bool bIsLastSuccessRopData, BookmarkType queryRowOrigin, bool isRequestCountTooLarger);

    /// <summary>
    /// This delegate is used to handle response which ResetTableRequest returned
    /// </summary>
    /// <param name="ropType">Identify rop type</param>
    /// <param name="isPreviousDataReset">Identify whether the previous data is reset</param>
    public delegate void RopResetTableHandler(TableRopType ropType, bool isPreviousDataReset);

    /// <summary>
    /// This delegate is used to handle response which RopQueryPositionRequest returned
    /// </summary>
    /// <param name="cursorPosition">Identify the current cursor position</param>
    /// <param name="isRowCount">Identify whether the right rows count is returned</param>
    public delegate void RopQueryPositionHandler(CursorPosition cursorPosition, bool isRowCount);

    /// <summary>
    /// This delegate is used to handle response which RopSeekRowRequest returned
    /// </summary>
    /// <param name="wantRowMovedCountout">Identify whether to return the actual seek row count</param>
    /// <param name="enoughRow">Identify whether there are enough row to seek</param>
    /// <param name="hasSoughtLess">Identify whether the actual seek row is less the one requested</param>
    /// <param name="isActualRowsCount">Identify whether the actual seek row is correct</param>
    public delegate void RopSeekRowHandler(bool wantRowMovedCountout, bool enoughRow, bool hasSoughtLess, bool isActualRowsCount);

    /// <summary>
    /// This delegate is used to handle response which RopSeekRowBookmarkRequest returned
    /// </summary>
    /// <param name="wantRowMovedCountout">Identify whether to return the actual seek row count</param>
    /// <param name="enoughRow">Identify whether there are enough row to seek</param>
    /// <param name="hasSoughtLess">Identify whether the actual seek row is less the one requested</param>
    /// <param name="isActualRowsCount">Identify whether the actual seek row is correct</param>
    /// <param name="rowNoLongerVisible">Identify whether the row is visible</param>
    public delegate void RopSeekRowBookmarkHandler(bool wantRowMovedCountout, bool enoughRow, bool hasSoughtLess, bool isActualRowsCount, bool rowNoLongerVisible);

    /// <summary>
    /// This delegate is used to handle response which RopSeekRowFractionalRequest returned
    /// </summary>
    /// <param name="zeroNumerator">Identify whether the Numerator is zero</param>
    /// <param name="smallDenominator">Identify whether the Numerator bigger than Denominator</param>
    /// <param name="cursorPosition">Cursor position</param>
    public delegate void RopSeekRowFractionalHandler(bool zeroNumerator, bool smallDenominator, CursorPosition cursorPosition);
    
    /// <summary>
    /// This delegate is used to handle response which RopFindRowRequest returned
    /// </summary>
    /// <param name="findRowFlags">Identify the flags for this operation</param>
    /// <param name="isValidRestriction">Identify whether the restriction is valid</param>
    /// <param name="HasRowData">Identify whether there are rows meet the restriction </param>
    /// <param name="isCorrectRowData">Identify whether row returned is correct</param>
    /// <param name="rowNoLongerVisible">Identify whether the row is visible</param>
    /// <param name="isLastSetColumnData">Identify whether the column data returned is the one set in the SetColumn method</param>
    public delegate void RopFindRowHandler(FindRowFlags findRowFlags, bool isValidRestriction, bool HasRowData, bool isCorrectRowData, bool rowNoLongerVisible, bool isLastSetColumnData);

    /// <summary>
    /// This delegate is used to handle response which RopExpandRowRequest returned
    /// </summary>
    /// <param name="isLastSetColumnData">Identify whether the column data returned is the one set in the SetColumn method</param>
    public delegate void RopExpandRowHandler(bool isLastSetColumnData);

    /// <summary>
    /// This delegate is used to handle response which RopGetCollapseStateRequest returned
    /// </summary>
    /// <param name="isSetCollapseStateCalled">Identify whether the SetCollapseState Rop is called</param>
    /// <param name="isCollapseStateRebuildCorrectly">Identify whether the collapse state is rebuilt successfully</param>
    public delegate void RopGetCollapseStateHandler(bool isSetCollapseStateCalled, bool isCollapseStateRebuildCorrectly);
    #endregion

    /// <summary>
    /// This interface defines the protocol adapter: IMS_OXCTABLAdapter
    /// </summary>
    public interface IMS_OXCTABLAdapter : IAdapter
    {
        #region Protocol event

        /// <summary>
        /// This event is used to handle response which RopQueryRowsRequest returned
        /// </summary>
        event RopQueryRowsHandler RopQueryRowsResponse;

        /// <summary>
        /// This event is used to handle response which RopResetTableRequest returned
        /// </summary>
        event RopResetTableHandler RopResetTableResponse;

        /// <summary>
        /// This event is used to handle response which RopQueryPositionRequest returned
        /// </summary>
        event RopQueryPositionHandler RopQueryPositionResponse;

        /// <summary>
        /// This event is used to handle response which RopSeekRowRequest returned
        /// </summary>
        event RopSeekRowHandler RopSeekRowResponse;

        /// <summary>
        /// This event is used to handle response which RopSeekRowBookmarkRequest returned
        /// </summary>
        event RopSeekRowBookmarkHandler RopSeekRowBookmarkResponse;

        /// <summary>
        /// This event is used to handle response which RopSeekRowFractionalRequest returned
        /// </summary>
        event RopSeekRowFractionalHandler RopSeekRowFractionalResponse;

        /// <summary>
        /// This event is used to handle response which RopFindRowRequest returned
        /// </summary>
        event RopFindRowHandler RopFindRowResponse;

        /// <summary>
        /// This event is used to handle response which RopExpandRowRequest returned
        /// </summary>
        event RopExpandRowHandler RopExpandRowResponse;

        /// <summary>
        /// This event is used to handle response which RopGetCollapseStateRequest returned
        /// </summary>
        event RopGetCollapseStateHandler RopGetCollapseStateResponse;
        #endregion

        #region Protocol operations
        /// <summary>
        /// This method is used to set columns for a table
        /// </summary>
        /// <param name="reqId">Identify the request ID</param>
        /// <param name="invalidPropertyTag">Indicate whether there are invalid property tags in the column set </param>
        /// <param name="newColumnSet">Indicate whether to change the current column set with new one</param>
        /// <param name="isSetColumnsAsynchronous">Indicate whether RopsetColumns Rop is to be performed asynchronously</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopSetColumns(uint reqId, bool invalidPropertyTag, bool newColumnSet, bool isSetColumnsAsynchronous);

        /// <summary>
        /// This method is used to sort table
        /// </summary>
        /// <param name="reqId">Identify the request ID</param>
        /// <param name="validHandle">Identify whether the InputHandleIndex is valid, this is used to trigger the error code in SortTable</param>
        /// <param name="multipleSortOrders">Indicate whether the sort order array contains more than one SortOrder structure</param>
        /// <param name="isMaximumCategory">Indicate whether the sort order array contains a SortOrder structure with the order set to MaximumCategory</param>
        /// <param name="allSortOrdersUsedAsCategory">Identify whether all sort orders used as category</param>
        /// <param name="allCategoryExpanded">Identify whether all categories are expanded</param>
        /// <param name="newOrder">Identify whether to change the current sort order</param>
        /// <param name="isSortTableAsynchronous">Indicate whether RopSortTable Rop is to be performed asynchronously</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopSortTable(uint reqId, bool validHandle, bool multipleSortOrders, bool isMaximumCategory, bool allSortOrdersUsedAsCategory, bool allCategoryExpanded, bool newOrder, bool isSortTableAsynchronous);

        /// <summary>
        /// This method is used to set restrict for a table
        /// </summary>
        /// <param name="reqId">Identify the request ID</param>
        /// <param name="validHandle">Identify whether the InputHandleIndex is valid, this is used to trigger the error code in Restrict</param>
        /// <param name="newRestrict">A flag identify whether new restriction is required</param>
        /// <param name="isRestrictAsynchronous">Indicate whether RopRestrict Rop is to be performed asynchronously</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopRestrict(uint reqId, bool validHandle, bool newRestrict, bool isRestrictAsynchronous);

        /// <summary>
        /// This method is used to query rows for a table
        /// </summary>
        /// <param name="queryRowFlags">A flag for the query row operation</param>
        /// <param name="forwardRead">Identify whether to read the data forward</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopQueryRows(QueryRowsFlags queryRowFlags, bool forwardRead);

        /// <summary>
        /// This method is used to Abort Rops for a table
        /// </summary>
        /// <param name="tableStatus">Identify the table status</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopAbort(out TableStatus tableStatus);

        /// <summary>
        /// This method is used to get the status for a table
        /// </summary>
        /// <param name="tableStatus">Identify the table status</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopGetStatus(out TableStatus tableStatus);

        /// <summary>
        /// This method is used to query position for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopQueryPosition();

        /// <summary>
        /// This method is used to seek row for a table
        /// </summary>
        /// <param name="origin">Identify the origin position</param>
        /// <param name="forwardSeek">Identify whether to seek row forward</param>
        /// <param name="wantRowMovedCountout">Identify whether to return the actual moved row number</param>
        /// <param name="enoughRow">Identify whether there is enough row to seek</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopSeekRow(BookmarkType origin, bool forwardSeek, bool wantRowMovedCountout, bool enoughRow);

        /// <summary>
        /// This method is used to seek row bookmark for a table
        /// </summary>
        /// <param name="forwardSeek">Identify whether to seek bookmark forward</param>
        /// <param name="wantRowMovedCountout">Identify whether to return the actual moved row number</param>
        /// <param name="enoughRow">Identify whether there is enough row to seek</param>
        /// <returns>Table rop return values</returns>
        TableRopReturnValues RopSeekRowBookmark(bool forwardSeek, bool wantRowMovedCountout, bool enoughRow);

        /// <summary>
        /// This method is used to seek row fractional for a table
        /// </summary>
        /// <param name="zeroNumerator">Identify whether the Numerator is zero</param>
        /// <param name="smallDenominator">Identify whether the Numerator bigger than Denominator</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopSeekRowFractional(bool zeroNumerator, bool smallDenominator);

        /// <summary>
        /// This method is used to create bookmark for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopCreateBookmark();

        /// <summary>
        /// This method is used to query all columns for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopQueryColumnsAll();

        /// <summary>
        /// This method is used to find row for a table
        /// </summary>
        /// <param name="findRowFlags">Identify the flags for this operation</param>
        /// <param name="origin">Identify the bookmark</param>
        /// <param name="isValidRestriction">Identify whether the restriction is valid</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopFindRow(FindRowFlags findRowFlags, BookmarkType origin, bool isValidRestriction);

        /// <summary>
        /// This method is used to free bookmark for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopFreeBookmark();

        /// <summary>
        /// This method is used to reset a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopResetTable();

        /// <summary>
        /// This method is used to expand row for a table
        /// </summary>
        /// <param name="nonexistCategoryId">Identify whether the CategoryId row exists</param>
        /// <param name="noncollapsedCategoryId">Identify whether the CategoryId row is not collapsed</param>
        /// <param name="isMaxRowCountInvalid">Identify whether the MaxRowCount is invalid</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopExpandRow(bool nonexistCategoryId, bool noncollapsedCategoryId, bool isMaxRowCountInvalid);

        /// <summary>
        /// This method is used to Collapse row for a table
        /// </summary>
        /// <param name="nonexistCategoryId">Identify whether the CategoryId row exists</param>
        /// <param name="nonexpandedCategoryId">Identify whether the CategoryId row is not expanded</param>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopCollapseRow(bool nonexistCategoryId, bool nonexpandedCategoryId);

        /// <summary>
        /// This method is used to get collapse state for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopGetCollapseState();

        /// <summary>
        /// This method is used to set collapse state for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        TableRopReturnValues RopSetCollapseState();

        /// <summary>
        /// This method is used to release a table
        /// </summary>
        void RopRelease();
        #endregion

        /// <summary>
        /// This method is used to initialize the table
        /// </summary>
        /// <param name="tableType">Identify the table type</param>
        void InitializeTable(TableType tableType);

        /// <summary>
        /// Determines if the requirement is enabled or not.
        /// </summary>
        /// <param name="requirementId">The requirement Id</param>
        /// <param name="enabled">Requirement is enable or not.</param>
        void CheckRequirementEnabled(int requirementId, out bool enabled);

        /// <summary>
        /// This method is used to check whether MAPIHTTP transport is supported by SUT.
        /// </summary>
        /// <param name="isSupported">The transport is supported or not.</param>
        void CheckMAPIHTTPTransportSupported(out bool isSupported);

        /// <summary>
        /// Disconnect from the server.
        /// </summary>
        void Disconnect();
    }
}