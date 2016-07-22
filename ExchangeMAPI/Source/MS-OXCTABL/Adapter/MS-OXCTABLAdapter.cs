namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implementation class of interface IOXCTABLAdapter
    /// </summary>
    public partial class MS_OXCTABLAdapter : ManagedAdapterBase, IMS_OXCTABLAdapter
    {
        #region Private variables

        /// <summary>
        /// The OxcropsClient instance.
        /// </summary>
        private OxcropsClient oxcropsClient;

        /// <summary>
        /// Status of connection.
        /// </summary>
        private bool isConnected;

        /// <summary>
        /// Raw data in the response
        /// </summary>
        private byte[] rawData;

        /// <summary>
        /// Input obj handle for the rop request
        /// </summary>
        private uint inputObjHandle;

        /// <summary>
        /// Identify of table Handle
        /// </summary>
        private uint tableHandle;

        /// <summary>
        /// Identify the type of table
        /// </summary>
        private TableType tableType;

        /// <summary>
        /// Identify of folder Handle
        /// </summary>
        private uint folderHandle;

        /// <summary>
        /// Identify the folder ID
        /// </summary>
        private ulong folderID;

        /// <summary>
        /// Rop response
        /// </summary>
        private object response;

        /// <summary>
        /// The list of out handles in the rop response
        /// </summary>
        private List<List<uint>> responseSOHs;

        /// <summary>
        /// The field used in RopSetColumns.
        /// </summary>
        private bool sign;

        /// <summary>
        /// Inbox folder's handle
        /// </summary>
        private uint inboxFolderHandle;

        /// <summary>
        /// The custom bookmark
        /// </summary>
        private byte[] userDefinedBookmark;

        /// <summary>
        /// Specifies the size of the userDefinedBookmark field
        /// </summary>
        private ushort userDefinedBookmarkSize;

        /// <summary>
        /// Identify the user defined bookmark position
        /// </summary>
        private uint userDefinedBookmarkPosition;

        /// <summary>
        /// Identify the property value
        /// </summary>
        private List<PropertyValue> propertyValues;

        /// <summary>
        /// This value specifies the collapse state for a categorized table is got.
        /// </summary>
        private bool isCollapseStateGot;

        /// <summary>
        /// This value specifies the collapse state is rebuilt.
        /// </summary>
        private bool isSetCollapseStateCalled;

        /// <summary>
        /// This array specifies a collapse state for a categorized table.
        /// </summary>
        private byte[] collapseState;

        /// <summary>
        /// This value specifies the size of the CollapseState field.
        /// </summary>
        private ushort collapseStateSize;

        /// <summary>
        /// Identify the row is Expanded.
        /// </summary>
        private bool isExpanded;

        /// <summary>
        /// Identify whether the sort order array contains more than one SortOrder structure.
        /// </summary>
        private bool areMultipleSortOrders;

        /// <summary>
        /// Identify whether all sort orders used as category
        /// </summary>
        private bool areAllSortOrdersUsedAsCategory;

        /// <summary>
        /// Identify whether all categories are used to expanded
        /// </summary>
        private bool areAllCategoryExpanded;

        /// <summary>
        /// Identify the number of rows in the table
        /// </summary>
        private uint rowCount;

        /// <summary>
        /// Identify the type of Rops operate
        /// </summary>
        private TableRopType ropType;

        /// <summary>
        /// Identify the propertyTags of last request
        /// </summary>
        private PropertyTag[] latestPropertyTags;

        /// <summary>
        /// Identify the propertyTags of last and success request
        /// </summary>
        private PropertyTag[] latestSuccessPropertyTages;

        /// <summary>
        /// Identify the order when latest SortTable
        /// </summary>
        private SortOrderFlag latestSortOrder;

        /// <summary>
        /// Identify the order when latest successful SortTable
        /// </summary>
        private SortOrderFlag latestSuccessSortOrder;

        /// <summary>
        /// Identify the restrictType of last request
        /// </summary>
        private RestrictFlag latestRestrict;

        /// <summary>
        /// Identify the restrictType of last and success request
        /// </summary>
        private RestrictFlag latestSuccessRestrict;

        /// <summary>
        /// Global field identify if SetColunms asynchronous
        /// </summary>
        private bool globalIsSetColumnsAsynchronous;

        /// <summary>
        /// Global field identify if SortTable asynchronous
        /// </summary>
        private bool globalIsSortTableAsynchronous;

        /// <summary>
        /// Global field identify if Restrict asynchronous
        /// </summary>
        private bool globalIsRestrictAsynchronous;

        /// <summary>
        /// Global field identify the current position in query rows response
        /// </summary>
        private BookmarkType queryRowOrigin;

        /// <summary>
        /// The value of the MaxRowCount field in the ROP request buffer(ExpandRow).
        /// </summary>
        private ushort maxRowCountInExpandRowRequest;

        /// <summary>
        /// The RowCount that is specified in the request(QueryRows).
        /// </summary>
        private ushort rowCountInQueryRowsRequest;

        /// <summary>
        /// The value of the RowCount field in the ROP request buffer(SeekRow).
        /// </summary>
        private int rowCountInSeekRowRequest;

        /// <summary>
        /// The value of the RowCount field in the ROP request buffer(SeekRowBookmark).
        /// </summary>
        private int rowCountInSeekRowBookmarkRequest;

        /// <summary>
        /// The time internal (millisecond) that is used by the test suite to find the message item after it sends out.
        /// </summary>
        private int waitTime;

        /// <summary>
        /// Indicate whether needs to do the cleanup work
        /// </summary>
        private bool needDoCleanup = true;
        #endregion

        #region IOXCTABLAdapter members
        
        #region Protocol event

        /// <summary>
        /// This event is used to handle response which RopQueryRowsRequest returned
        /// </summary>
        public event RopQueryRowsHandler RopQueryRowsResponse;

        /// <summary>
        /// This event is used to handle response which RopResetTableRequest returned
        /// </summary>
        public event RopResetTableHandler RopResetTableResponse;

        /// <summary>
        /// This event is used to handle response which RopQueryPositionRequest returned
        /// </summary>
        public event RopQueryPositionHandler RopQueryPositionResponse;

        /// <summary>
        /// This event is used to handle response which RopSeekRowRequest returned
        /// </summary>
        public event RopSeekRowHandler RopSeekRowResponse;

        /// <summary>
        /// This event is used to handle response which RopSeekRowBookmarkRequest returned
        /// </summary>
        public event RopSeekRowBookmarkHandler RopSeekRowBookmarkResponse;

        /// <summary>
        /// This event is used to handle response which RopSeekRowFractionalRequest returned
        /// </summary>
        public event RopSeekRowFractionalHandler RopSeekRowFractionalResponse;

        /// <summary>
        /// This event is used to handle response which RopFindRowRequest returned
        /// </summary>
        public event RopFindRowHandler RopFindRowResponse;

        /// <summary>
        /// This event is used to handle response which RopExpandRowRequest returned
        /// </summary>
        public event RopExpandRowHandler RopExpandRowResponse;

        /// <summary>
        /// This event is used to handle response which RopGetCollapseStateRequest returned
        /// </summary>
        public event RopGetCollapseStateHandler RopGetCollapseStateResponse;

        #endregion

        #region Protocol operations
        /// <summary>
        /// This method is used to SetColumns for a table
        /// </summary>
        /// <param name="reqId">Identify the request ID</param>
        /// <param name="invalidPropertyTag">Indicate whether there are invalid property tags in the column set </param>
        /// <param name="newColumnSet">Indicate whether to change the current column set with new one</param>
        /// <param name="isSetColumnsAsynchronous">Indicate whether RopsetColumns Rop is to be performed asynchronously</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopSetColumns(uint reqId, bool invalidPropertyTag, bool newColumnSet, bool isSetColumnsAsynchronous)
        {
            RopSetColumnsRequest setColumnsRequest;
            RopSetColumnsResponse setColumnsResponse;
            PropertyTag[] propertyTags;

            if (invalidPropertyTag)
            {
                propertyTags = this.CreateInvalidContentsTablePropertyTags();
            }
            else
            {
                switch (this.tableType)
                {
                    case TableType.CONTENT_TABLE:
                        if (!newColumnSet && this.sign)
                        {
                            propertyTags = this.CreateSampleContentsTablePropertyTags();
                        }
                        else
                        {
                            propertyTags = this.CreateSampleContentsTablePropertyTagsWithNoError();
                        }

                        break;
                    case TableType.HIERARCHY_TABLE:
                        propertyTags = this.CreateSampleHierarchyTablePropertyTags();
                        break;
                    case TableType.RULES_TABLE:
                        propertyTags = this.CreateRulesTablePropertyTags();
                        break;
                    default:
                        propertyTags = this.CreateSampleContentsTablePropertyTags();
                        break;
                }

                this.sign = this.sign == !newColumnSet;
            }

            setColumnsRequest.RopId = 0x12;
            setColumnsRequest.LogonId = 0x00;
            setColumnsRequest.InputHandleIndex = 0x00;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;
            if (isSetColumnsAsynchronous)
            {
                setColumnsRequest.SetColumnsFlags = 0x01; // Async
            }
            else
            {
                setColumnsRequest.SetColumnsFlags = 0x00; // Sync
            }

            this.globalIsSetColumnsAsynchronous = isSetColumnsAsynchronous;
            this.DoSingleCallROP(setColumnsRequest, this.tableHandle, ref this.response, ref this.rawData);
            setColumnsResponse = (RopSetColumnsResponse)this.response;

            this.ropType = TableRopType.SETCOLUMNS;
            this.latestPropertyTags = propertyTags;
            if ((TableRopReturnValues)setColumnsResponse.ReturnValue == TableRopReturnValues.success)
            {
                this.latestSuccessPropertyTages = propertyTags;
            }

            // If the RopSetColumns ROP is not complete (the returned TableStatus is set to 0x0B), wait for the RopGetStatus ROP to return TableStatus set 0x00 (which indicates the RopSetColumns ROP is complete).
            if ((TableRopReturnValues)setColumnsResponse.ReturnValue == TableRopReturnValues.success && isSetColumnsAsynchronous && (TableStatus)setColumnsResponse.TableStatus == TableStatus.TblstatSettingCols)
            {
                int getStatusRetryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                RopGetStatusResponse getStatusResponse = new RopGetStatusResponse();
                while (getStatusRetryCount >= 0)
                {
                    Thread.Sleep(this.waitTime);
                    TableStatus tableStatus;
                    this.RopGetStatus(out tableStatus);
                    if (tableStatus == TableStatus.TblstatComplete)
                    {
                        break;
                    }

                    getStatusRetryCount--;
                }

                if (getStatusRetryCount < 0)
                {
                    Site.Assert.Fail("The RopSetColumns ROP should complete!");
                }

                this.VerifyAsynchronousROPComplete(getStatusResponse);
            }

            return (TableRopReturnValues)setColumnsResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to sort table
        /// </summary>
        /// <param name="reqId">Identify the request ID</param>
        /// <param name="validHandle">Identify whether the InputHandleIndex is valid, this is used to trigger the error code in sort table</param>
        /// <param name="multipleSortOrders">Indicate whether the sort order array contains more than one SortOrder structure</param>
        /// <param name="isMaximumCategory">Indicate whether the sort order array contains a SortOrder structure with the order set to MaximumCategory</param>
        /// <param name="allSortOrdersUsedAsCategory">Identify whether all sort orders used as category</param>
        /// <param name="allCategoryExpanded">Identify whether all categories are expanded</param>
        /// <param name="newOrder">Identify whether to change the current sort order</param>
        /// <param name="isSortTableAsynchronous">Indicate whether RopSortTable Rop is to be performed asynchronously</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopSortTable(uint reqId, bool validHandle, bool multipleSortOrders, bool isMaximumCategory, bool allSortOrdersUsedAsCategory, bool allCategoryExpanded, bool newOrder, bool isSortTableAsynchronous)
        {
            RopSortTableRequest sortTableRequest;
            RopSortTableResponse sortTableResponse;
            SortOrder[] sortOrders;

            if (!newOrder && this.sign)
            {
                sortOrders = this.CreateSampleSortOrdersAscending(multipleSortOrders, isMaximumCategory);
            }
            else
            {
                sortOrders = this.CreateSampleSortOrdersDescending(multipleSortOrders);
            }

            this.sign = this.sign == !newOrder;
            this.latestSortOrder = sortOrders[0].Order == 0x00 ? SortOrderFlag.SortOrderASC : SortOrderFlag.SortOrderDESC;
            sortTableRequest.RopId = 0x13;
            sortTableRequest.LogonId = 0x00;
            sortTableRequest.InputHandleIndex = (byte)0x00; 
            this.areAllSortOrdersUsedAsCategory = allSortOrdersUsedAsCategory;
            this.areAllCategoryExpanded = allCategoryExpanded;
            this.areMultipleSortOrders = multipleSortOrders;
            sortTableRequest.SortOrderCount = (ushort)sortOrders.Length;
            if (allSortOrdersUsedAsCategory)
            {
                sortTableRequest.CategoryCount = (ushort)sortOrders.Length;
                if (allCategoryExpanded)
                {
                    sortTableRequest.ExpandedCount = (ushort)sortOrders.Length; // All rows expanded
                }
                else
                {
                    sortTableRequest.ExpandedCount = (ushort)(sortOrders.Length - 1); // Not all rows expanded
                }
            }
            else
            {
                sortTableRequest.CategoryCount = (ushort)(sortOrders.Length - 1);
                if (allCategoryExpanded)
                {
                    sortTableRequest.ExpandedCount = (ushort)(sortOrders.Length - 1); // All rows expanded
                }
                else
                {
                    sortTableRequest.ExpandedCount = 0;
                }
            }

            sortTableRequest.SortOrders = sortOrders;
            if (isSortTableAsynchronous)
            {
                sortTableRequest.SortTableFlags = 0x01; // Async
            }
            else
            {
                sortTableRequest.SortTableFlags = 0x00; // Sync
            }

            uint inputObjHandle;
            if (validHandle)
            {
                inputObjHandle = this.tableHandle;
            }
            else
            {
                inputObjHandle = 0xFFFFFFFF; // Cause the request failed when set to invalid handle
            }

            this.globalIsSortTableAsynchronous = isSortTableAsynchronous;
            this.DoSingleCallROP(sortTableRequest, inputObjHandle, ref this.response, ref this.rawData);
            sortTableResponse = (RopSortTableResponse)this.response;

            this.ropType = TableRopType.SORTTABLE;

            if ((TableRopReturnValues)sortTableResponse.ReturnValue == TableRopReturnValues.success)
            {
                this.latestSuccessSortOrder = this.latestSortOrder;
                if (sortTableRequest.ExpandedCount != 0)
                {
                    this.isExpanded = true; // The first header row is expanded when the request succeeds.
                }

                this.GetRowsCount(false); // Update the field rowCount
            }

            // If the RopSortTable ROP is not complete (the returned TableStatus is set to 0x09), wait for the RopGetStatus ROP to return TableStatus set 0x00 (which indicates the RopSortTable ROP is complete).
            if ((TableRopReturnValues)sortTableResponse.ReturnValue == TableRopReturnValues.success && isSortTableAsynchronous && (TableStatus)sortTableResponse.TableStatus == TableStatus.TblstatSorting)
            {
                int getStatusRetryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                RopGetStatusResponse getStatusResponse = new RopGetStatusResponse();
                while (getStatusRetryCount >= 0)
                {
                    Thread.Sleep(this.waitTime);
                    TableStatus tableStatus;
                    this.RopGetStatus(out tableStatus);
                    if (tableStatus == TableStatus.TblstatComplete)
                    {
                        break;
                    }

                    getStatusRetryCount--;
                }

                if (getStatusRetryCount < 0)
                {
                    Site.Assert.Fail("The RopSortTable ROP should complete!");
                }

                this.VerifyAsynchronousROPComplete(getStatusResponse);
            }

            return (TableRopReturnValues)sortTableResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to restrict for a table
        /// </summary>
        /// <param name="reqId">Identify the request ID</param>
        /// <param name="validHandle">Identify whether the InputHandleIndex is valid, this is used to trigger the error code in Restrict</param>
        /// <param name="newRestrict">A identify whether new restriction is required</param>
        /// <param name="isRestrictAsynchronous">Indicate whether RopRestrict Rop is to be performed asynchronously</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopRestrict(uint reqId, bool validHandle, bool newRestrict, bool isRestrictAsynchronous)
        {
            RopRestrictRequest restrictRequest;
            RopRestrictResponse restrictResponse;

            restrictRequest.RopId = 0x14;
            restrictRequest.LogonId = 0x00;
            restrictRequest.InputHandleIndex = (byte)0x00; 
            if (isRestrictAsynchronous)
            {
                restrictRequest.RestrictFlags = 0x01; // Async
            }
            else
            {
                restrictRequest.RestrictFlags = 0x00; // Sync
            }

            this.globalIsRestrictAsynchronous = isRestrictAsynchronous;
            byte[] data;
            if (this.tableType == TableType.CONTENT_TABLE)
            {
                if (!newRestrict && this.sign)
                {
                    byte[] tempData = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("Sender1Name", this.Site) + "\0");
                    data = new byte[tempData.Length + 13];
                    int index = 0;
                    data[index++] = 0x03;
                    data[index++] = 0x02;
                    data[index++] = 0x00;
                    data[index++] = 0x04;
                    data[index++] = 0x00;
                    data[index++] = 0x1f;
                    data[index++] = 0x00;
                    data[index++] = 0x1a;
                    data[index++] = 0x0c;
                    data[index++] = 0x1f;
                    data[index++] = 0x00;
                    data[index++] = 0x1a;
                    data[index++] = 0x0c;
                    Array.Copy(tempData, 0, data, index, tempData.Length);

                    this.latestRestrict = RestrictFlag.SenderIsTest1Restrict;
                }
                else
                {
                    byte[] tempData = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("Sender2Name", this.Site) + "\0");
                    data = new byte[tempData.Length + 13];
                    int index = 0;
                    data[index++] = 0x03;
                    data[index++] = 0x02;
                    data[index++] = 0x00;
                    data[index++] = 0x04;
                    data[index++] = 0x00;
                    data[index++] = 0x1f;
                    data[index++] = 0x00;
                    data[index++] = 0x1a;
                    data[index++] = 0x0c;
                    data[index++] = 0x1f;
                    data[index++] = 0x00;
                    data[index++] = 0x1a;
                    data[index++] = 0x0c;
                    Array.Copy(tempData, 0, data, index, tempData.Length);

                    this.latestRestrict = RestrictFlag.SenderIsTest2Restriction;
                }

                this.sign = this.sign == !newRestrict;
            }
            else if (this.tableType == TableType.HIERARCHY_TABLE || this.tableType == TableType.RULES_TABLE)
            {
                data = AdapterHelper.GenerateValidRestrictData(this.Site, this.tableType);
            }
            else
            {
                data = null;
            }

            restrictRequest.RestrictionDataSize = data == null ? (ushort)0 : (ushort)data.Length;
            restrictRequest.RestrictionData = data;

            uint inputObjHandle;
            if (validHandle)
            {
                inputObjHandle = this.tableHandle;
            }
            else
            {
                inputObjHandle = 0xFFFFFFFF; // Cause the request failed when set to invalid handle
            }

            this.DoSingleCallROP(restrictRequest, inputObjHandle, ref this.response, ref this.rawData);
            restrictResponse = (RopRestrictResponse)this.response;

            this.ropType = TableRopType.RESTRICT;

            if ((TableRopReturnValues)restrictResponse.ReturnValue == TableRopReturnValues.success)
            {
                this.latestSuccessRestrict = this.latestRestrict;

                this.GetRowsCount(false);
            }

            // If the RopRestrict ROP is not complete (the returned TableStatus is set to 0x0E), wait for the RopGetStatus ROP to return TableStatus set 0x00 (which indicates the RopRestrict ROP is complete).
            if ((TableRopReturnValues)restrictResponse.ReturnValue == TableRopReturnValues.success && isRestrictAsynchronous && (TableStatus)restrictResponse.TableStatus == TableStatus.TblstatRestricting)
            {
                int getStatusRetryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                RopGetStatusResponse getStatusResponse = new RopGetStatusResponse();
                while (getStatusRetryCount >= 0)
                {
                    Thread.Sleep(this.waitTime);
                    TableStatus tableStatus;
                    this.RopGetStatus(out tableStatus);
                    if (tableStatus == TableStatus.TblstatComplete)
                    {
                        break;
                    }

                    getStatusRetryCount--;
                }

                if (getStatusRetryCount < 0)
                {
                    Site.Assert.Fail("The RopRestrict ROP should complete!");
                }

                this.VerifyAsynchronousROPComplete(getStatusResponse);
            }

            return (TableRopReturnValues)restrictResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to query rows for a table
        /// </summary>
        /// <param name="queryRowFlags">A flag for the query row operation</param>
        /// <param name="forwardRead">Identify whether to read the data forward</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopQueryRows(QueryRowsFlags queryRowFlags, bool forwardRead)
        {
            // The enabled property value for MS-OXCTABL_R768, MS-OXCTABL_R866 and MS-OXCTABL_R867 must the same.
            if (!(Common.IsRequirementEnabled(768, this.Site) == Common.IsRequirementEnabled(866, this.Site) == Common.IsRequirementEnabled(867, this.Site) == true) && !(Common.IsRequirementEnabled(768, this.Site) == Common.IsRequirementEnabled(866, this.Site) == Common.IsRequirementEnabled(867, this.Site) == false))
            {
                Site.Assert.Inconclusive("Requirements implementation conflict, the conflict requirements number are: MS-OXCTABL_R768, MS-OXCTABL_R866 and MS-OXCTABL_R867");
            }

            uint tempRowCount = this.GetRowsCount(forwardRead); // Gets the number of rows from current location to begin row or last row.

            RopQueryRowsRequest queryRowsRequest;
            RopQueryRowsResponse queryRowsResponse;

            queryRowsRequest.RopId = 0x15;
            queryRowsRequest.LogonId = 0x00;
            queryRowsRequest.InputHandleIndex = 0x00;
            queryRowsRequest.QueryRowsFlags = (byte)queryRowFlags; // Whether advance the table cursor. 
            queryRowsRequest.ForwardRead = Convert.ToByte(forwardRead); // It is set to FALSE (0x00) to read the table backwards. It MUST NOT be set to any other value.
            queryRowsRequest.RowCount = 0x0003; // Return a maximum of 3 rows.
            this.rowCountInQueryRowsRequest = queryRowsRequest.RowCount;

            if (Common.IsRequirementEnabled(768, this.Site))
            {
                this.DoSingleCallROP(queryRowsRequest, this.tableHandle, ref this.response, ref this.rawData);
            }
            else
            {
                // In order to test the case when buffer is smaller.
                try
                {
                    this.DoSingleCallROP(queryRowsRequest, this.tableHandle, ref this.response, ref this.rawData);
                }
                catch (Exception ex)
                {
                    if (ex.Message == Constants.BufferTooSmall)
                    {
                        return TableRopReturnValues.ecBufferTooSmall;
                    }
                    else
                    {
                        Site.Log.Add(LogEntryKind.Comment, ex.Message);
                        Site.Log.Add(LogEntryKind.Comment, "The behavior when a RopQueryRows ROP is sent before a successful RopSetColumns ROP is undefined for implementation following Exchange 2007.");
                        return TableRopReturnValues.unexpected;
                    }
                }
            }

            queryRowsResponse = (RopQueryRowsResponse)this.response;

            int queryRowCount = 0;
            bool isLatestRopData = false;
            bool isLastSuccessRopData = false;
            bool zeroRow = false;
            bool isCorrectRowCountReturned = false;
            bool cursorPositionChanged = false;
            bool isRequestCountTooLarger = this.rowCountInQueryRowsRequest > tempRowCount;
            if ((TableRopReturnValues)queryRowsResponse.ReturnValue == TableRopReturnValues.success)
            {
                queryRowCount = queryRowsResponse.RowCount;
                this.queryRowOrigin = (BookmarkType)queryRowsResponse.Origin;
                isCorrectRowCountReturned = tempRowCount > this.rowCountInQueryRowsRequest ? queryRowCount == this.rowCountInQueryRowsRequest : queryRowCount == tempRowCount; // It MUST be less than or equal to the RowCount that is specified in the request
                cursorPositionChanged = tempRowCount != this.GetRowsCount(forwardRead);
                if (queryRowCount == 0)
                {
                    zeroRow = true;
                }
                else
                {
                    this.propertyValues = queryRowsResponse.RowData.PropertyRows[0].PropertyValues; // Gets the first row 
                    isLastSuccessRopData = this.latestSuccessPropertyTages.Length == queryRowsResponse.RowData.PropertyRows[0].PropertyValues.Count;
                    isLatestRopData = this.latestPropertyTags.Length == queryRowsResponse.RowData.PropertyRows[0].PropertyValues.Count;

                    if (this.tableType == TableType.CONTENT_TABLE)
                    {
                        bool isCorrectSortOrder = true;

                        // Gets the value of PidTagSenderName.
                        string tempSenderName = Encoding.Unicode.GetString(queryRowsResponse.RowData.PropertyRows[0].PropertyValues[2].Value);
                        if (this.latestSuccessSortOrder == SortOrderFlag.SortOrderASC)
                        {
                            // The PidTagSenderName must be Test1 when sortOrder is ascend.
                            isCorrectSortOrder = this.IsExpectedSender(tempSenderName, Common.GetConfigurationPropertyValue("Sender1Name", this.Site));
                        }
                        else if (this.latestSuccessSortOrder == SortOrderFlag.SortOrderDESC)
                        {
                            // The PidTagSenderName must be Test2 when sortOrder is descend.
                            isCorrectSortOrder = this.IsExpectedSender(tempSenderName, Common.GetConfigurationPropertyValue("Sender2Name", this.Site));
                        }

                        bool isCorrectRestriction = true;
                        if (this.latestSuccessRestrict == RestrictFlag.SenderIsTest1Restrict)
                        {
                            foreach (PropertyRow pr in queryRowsResponse.RowData.PropertyRows)
                            {
                                if (!this.IsExpectedSender(Encoding.Unicode.GetString(pr.PropertyValues[2].Value), Common.GetConfigurationPropertyValue("Sender1Name", this.Site)))
                                {
                                    isCorrectRestriction = false;
                                }
                            }
                        }

                        if (this.latestSuccessRestrict == RestrictFlag.SenderIsTest2Restriction)
                        {
                            foreach (PropertyRow pr in queryRowsResponse.RowData.PropertyRows)
                            {
                                if (!this.IsExpectedSender(Encoding.Unicode.GetString(pr.PropertyValues[2].Value), Common.GetConfigurationPropertyValue("Sender2Name", this.Site)))
                                {
                                    isCorrectRestriction = false;
                                }
                            }
                        }

                        bool isFirstHeaderExpanded = true;
                        bool areAllCategoryExpanded = true;
                        bool isCorrectSortOrderAndCategory = true;
                        if (this.latestSuccessSortOrder != SortOrderFlag.NotSort)
                        {
                            if (this.isExpanded == true)
                            {
                                if (BitConverter.ToInt32(queryRowsResponse.RowData.PropertyRows[0].PropertyValues[3].Value, 0) != 0x03)
                                {
                                    isFirstHeaderExpanded = false;
                                }
                            }

                            if (this.areAllCategoryExpanded)
                            {
                                foreach (PropertyRow pr in queryRowsResponse.RowData.PropertyRows)
                                {
                                    if (BitConverter.ToInt32(pr.PropertyValues[3].Value, 0) == 0x04)
                                    {
                                        areAllCategoryExpanded = false;
                                        break;
                                    }
                                }
                            }

                            if (!this.areAllSortOrdersUsedAsCategory)
                            {
                                if (this.areMultipleSortOrders)
                                {
                                    foreach (PropertyRow pr in queryRowsResponse.RowData.PropertyRows)
                                    {
                                        // Since the PidTagSenderName property is used in the first sortOrder structure, it will be used as category.
                                        if (BitConverter.ToInt32(pr.PropertyValues[3].Value, 0) == 0x03 || BitConverter.ToInt32(pr.PropertyValues[3].Value, 0) == 0x04)
                                        {
                                            if (!this.IsExpectedSender(Encoding.Unicode.GetString(pr.PropertyValues[2].Value), Common.GetConfigurationPropertyValue("Sender1Name", this.Site)) &&
                                                !this.IsExpectedSender(Encoding.Unicode.GetString(pr.PropertyValues[2].Value), Common.GetConfigurationPropertyValue("Sender2Name", this.Site)))
                                            {
                                                isCorrectSortOrderAndCategory = false;
                                                break;
                                            }
                                        }
                                    }

                                    if (isCorrectSortOrderAndCategory)
                                    {
                                        for (int i = 0; i < queryRowsResponse.RowData.PropertyRows.Count - 1; i++)
                                        {
                                            // Since the PidTagInstID property is used in the second sortOrder structure, it will be used to sort the rows within the category.
                                            if (BitConverter.ToInt32(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[3].Value, 0) == 0x01 &&
                                                BitConverter.ToInt32(queryRowsResponse.RowData.PropertyRows[i + 1].PropertyValues[3].Value, 0) == 0x01)
                                            {
                                                if ((this.latestSuccessSortOrder == SortOrderFlag.SortOrderASC && BitConverter.ToInt32(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[0].Value, 0) <
                                                    BitConverter.ToInt32(queryRowsResponse.RowData.PropertyRows[i + 1].PropertyValues[0].Value, 0)) ||
                                                    (this.latestSuccessSortOrder == SortOrderFlag.SortOrderDESC && BitConverter.ToInt32(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[0].Value, 0) >
                                                    BitConverter.ToInt32(queryRowsResponse.RowData.PropertyRows[i + 1].PropertyValues[0].Value, 0)))
                                                {
                                                    isCorrectSortOrderAndCategory = false;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        isLastSuccessRopData = isLastSuccessRopData && isCorrectSortOrder && isCorrectRestriction && isFirstHeaderExpanded && areAllCategoryExpanded && isCorrectSortOrderAndCategory;

                        if (this.latestSortOrder == SortOrderFlag.SortOrderASC)
                        {
                            // The PidTagSenderName must be Test1 when sortOrder is ascend.
                            isCorrectSortOrder = this.IsExpectedSender(tempSenderName, Common.GetConfigurationPropertyValue("Sender1Name", this.Site));
                        }
                        else if (this.latestSortOrder == SortOrderFlag.SortOrderDESC)
                        {
                            // The PidTagSenderName must be Test2 when sortOrder is descend.
                            isCorrectSortOrder = this.IsExpectedSender(tempSenderName, Common.GetConfigurationPropertyValue("Sender2Name", this.Site));
                        }

                        if (this.latestRestrict == RestrictFlag.SenderIsTest1Restrict)
                        {
                            foreach (PropertyRow pr in queryRowsResponse.RowData.PropertyRows)
                            {
                                if (!this.IsExpectedSender(Encoding.Unicode.GetString(pr.PropertyValues[2].Value), Common.GetConfigurationPropertyValue("Sender1Name", this.Site)))
                                {
                                    isCorrectRestriction = false;
                                }
                            }
                        }

                        if (this.latestRestrict == RestrictFlag.SenderIsTest2Restriction)
                        {
                            foreach (PropertyRow pr in queryRowsResponse.RowData.PropertyRows)
                            {
                                if (!this.IsExpectedSender(Encoding.Unicode.GetString(pr.PropertyValues[2].Value), Common.GetConfigurationPropertyValue("Sender2Name", this.Site)))
                                {
                                    isCorrectRestriction = false;
                                }
                            }
                        }

                        isLatestRopData = isLatestRopData && isCorrectSortOrder && isCorrectRestriction;
                    }
                }
            }
            else
            {
                isCorrectRowCountReturned = false;
            }

            if (this.RopQueryRowsResponse != null)
            {
                this.RopQueryRowsResponse(queryRowFlags, forwardRead, zeroRow, this.ropType, isCorrectRowCountReturned, cursorPositionChanged, isLatestRopData, isLastSuccessRopData, this.queryRowOrigin, isRequestCountTooLarger);
            }

            return (TableRopReturnValues)queryRowsResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to Abort Rops for a table
        /// </summary>
        /// <param name="tableStatus">Identify the table status</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopAbort(out TableStatus tableStatus)
        {
            RopAbortRequest abortRequest;
            RopAbortResponse abortResponse;

            abortRequest.RopId = 0x38;
            abortRequest.LogonId = 0x00;
            abortRequest.InputHandleIndex = 0x00;

            this.DoSingleCallROP(abortRequest, this.tableHandle, ref this.response, ref this.rawData);
            abortResponse = (RopAbortResponse)this.response;

            tableStatus = (TableStatus)abortResponse.TableStatus;

            return (TableRopReturnValues)abortResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to get the status for a table
        /// </summary>
        /// <param name="tableStatus">Identify the table status</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopGetStatus(out TableStatus tableStatus)
        {
            RopGetStatusRequest getStatusRequest;
            RopGetStatusResponse getStatusResponse;

            getStatusRequest.RopId = 0x16;
            getStatusRequest.LogonId = 0x00;
            getStatusRequest.InputHandleIndex = 0x00;

            this.DoSingleCallROP(getStatusRequest, this.tableHandle, ref this.response, ref this.rawData);
            getStatusResponse = (RopGetStatusResponse)this.response;

            tableStatus = (TableStatus)getStatusResponse.TableStatus;

            return (TableRopReturnValues)getStatusResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to query position for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopQueryPosition()
        {
            RopQueryPositionRequest queryPositionRequest;
            RopQueryPositionResponse queryPositionResponse;

            queryPositionRequest.RopId = 0x17;
            queryPositionRequest.LogonId = 0x00;
            queryPositionRequest.InputHandleIndex = 0x00;

            this.DoSingleCallROP(queryPositionRequest, this.tableHandle, ref this.response, ref this.rawData);
            queryPositionResponse = (RopQueryPositionResponse)this.response;

            uint numerator = queryPositionResponse.Numerator;
            uint denominator = queryPositionResponse.Denominator;

            CursorPosition currentPosition;
            if (numerator == 0)
            {
                // Indicate the cursor position is begin when numerator is 0.
                currentPosition = CursorPosition.BEGIN;
            }
            else if (numerator == denominator)
            {
                // Indicate the cursor position is end when numerator equal denominator.
                currentPosition = CursorPosition.END;
            }
            else
            {
                currentPosition = CursorPosition.CURRENT;
            }

            bool isRowCount = this.rowCount == denominator;

            if (this.RopQueryPositionResponse != null)
            {
                this.RopQueryPositionResponse(currentPosition, isRowCount);
            }

            return (TableRopReturnValues)queryPositionResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to seek row for a table
        /// </summary>
        /// <param name="origin">Identify the origin position</param>
        /// <param name="forwardSeek">Identify whether to seek row forward</param>
        /// <param name="wantRowMovedCountout">Identify whether to return the actual moved row number</param>
        /// <param name="enoughRow">Identify whether there is enough row to seek</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopSeekRow(BookmarkType origin, bool forwardSeek, bool wantRowMovedCountout, bool enoughRow)
        {
            uint tempRowCount = this.GetRowsCount(forwardSeek); // Gets the number of rows from current location to begin row or last row.

            RopSeekRowRequest seekRowRequest = new RopSeekRowRequest();
            RopSeekRowResponse seekRowResponse;

            seekRowRequest.RopId = 0x18;
            seekRowRequest.LogonId = 0x00;
            seekRowRequest.InputHandleIndex = 0x00;
            seekRowRequest.Origin = (byte)origin;

            if (origin == BookmarkType.BOOKMARK_BEGINNING)
            {
                if (forwardSeek)
                {
                    if (enoughRow)
                    {
                        seekRowRequest.RowCount = (int)this.rowCount;
                    }
                    else
                    {
                        // Identify there is not enough row to forward seek.
                        seekRowRequest.RowCount = (int)(this.rowCount + 1);
                    }
                }
                else
                {
                    if (enoughRow)
                    {
                        seekRowRequest.RowCount = 0;
                    }
                    else
                    {
                        seekRowRequest.RowCount = -1; // Identify there is not enough row to backward seek when the location is beginning.
                    }
                }
            }
            else if (origin == BookmarkType.BOOKMARK_END)
            {
                if (forwardSeek)
                {
                    if (enoughRow)
                    {
                        seekRowRequest.RowCount = 0;
                    }
                    else
                    {
                        seekRowRequest.RowCount = 1;
                    }
                }
                else
                {
                    if (enoughRow)
                    {
                        seekRowRequest.RowCount = -(int)this.rowCount;
                    }
                    else
                    {
                        seekRowRequest.RowCount = -(int)(this.rowCount + 1);
                    }
                }
            }
            else
            {
                if (forwardSeek)
                {
                    if (enoughRow)
                    {
                        seekRowRequest.RowCount = (int)tempRowCount;
                    }
                    else
                    {
                        seekRowRequest.RowCount = (int)(tempRowCount + 1);
                    }
                }
                else
                {
                    if (enoughRow)
                    {
                        seekRowRequest.RowCount = -(int)tempRowCount;
                    }
                    else
                    {
                        seekRowRequest.RowCount = -(int)(tempRowCount + 1);
                    }
                }
            }

            seekRowRequest.WantRowMovedCount = Convert.ToByte(wantRowMovedCountout);

            this.rowCountInSeekRowRequest = seekRowRequest.RowCount;

            this.DoSingleCallROP(seekRowRequest, this.tableHandle, ref this.response, ref this.rawData);
            seekRowResponse = (RopSeekRowResponse)this.response;

            this.ropType = TableRopType.MOVECURSOR;
            bool hasSoughtLess = false;
            bool isActualRowsCount = false;
            if (wantRowMovedCountout)
            {
                hasSoughtLess = Convert.ToBoolean(seekRowResponse.HasSoughtLess);
                if (origin == BookmarkType.BOOKMARK_BEGINNING)
                {
                    if (forwardSeek)
                    {
                        isActualRowsCount = seekRowResponse.RowsSought == this.rowCount;
                    }
                    else
                    {
                        isActualRowsCount = seekRowResponse.RowsSought == 0;
                    }
                }

                if (origin == BookmarkType.BOOKMARK_END)
                {
                    if (forwardSeek)
                    {
                        isActualRowsCount = seekRowResponse.RowsSought == 0;
                    }
                    else
                    {
                        isActualRowsCount = seekRowResponse.RowsSought == -this.rowCount;
                    }
                }

                if (origin == BookmarkType.BOOKMARK_CURRENT)
                {
                    isActualRowsCount = System.Math.Abs(seekRowResponse.RowsSought) == tempRowCount;
                }
            }
            else
            {
                hasSoughtLess = true;
                isActualRowsCount = true;
            }

            if (this.RopSeekRowResponse != null)
            {
                this.RopSeekRowResponse(wantRowMovedCountout, enoughRow, hasSoughtLess, isActualRowsCount);
            }

            return (TableRopReturnValues)seekRowResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to seek row bookmark for a table
        /// </summary>
        /// <param name="forwardSeek">Identify whether to seek bookmark forward</param>
        /// <param name="wantRowMovedCountout">Identify whether to return the actual moved row number</param>
        /// <param name="enoughRow">Identify whether there is enough row to seek</param>
        /// <returns>Table rop return values</returns>
        public TableRopReturnValues RopSeekRowBookmark(bool forwardSeek, bool wantRowMovedCountout, bool enoughRow)
        {
            RopSeekRowBookmarkRequest seekRowBookmarkRequest;
            RopSeekRowBookmarkResponse seekRowBookmarkResponse;

            seekRowBookmarkRequest.RopId = 0x19;
            seekRowBookmarkRequest.LogonId = 0x00;
            seekRowBookmarkRequest.InputHandleIndex = 0x00;
            seekRowBookmarkRequest.Bookmark = this.userDefinedBookmark; // Returned by a previous RopCreateBookmark request.
            seekRowBookmarkRequest.BookmarkSize = this.userDefinedBookmark == null ? (ushort)0 : (ushort)this.userDefinedBookmark.Length;
            if (forwardSeek)
            {
                if (enoughRow)
                {
                    seekRowBookmarkRequest.RowCount = (int)(this.rowCount - this.userDefinedBookmarkPosition);
                }
                else
                {
                    seekRowBookmarkRequest.RowCount = (int)this.rowCount + 1;
                }
            }
            else
            {
                if (enoughRow)
                {
                    seekRowBookmarkRequest.RowCount = (int)(-this.userDefinedBookmarkPosition);
                }
                else
                {
                    seekRowBookmarkRequest.RowCount = (int)(-(this.userDefinedBookmarkPosition + 1));
                }
            }

            seekRowBookmarkRequest.WantRowMovedCount = Convert.ToByte(wantRowMovedCountout);

            this.rowCountInSeekRowBookmarkRequest = seekRowBookmarkRequest.RowCount;

            this.DoSingleCallROP(seekRowBookmarkRequest, this.tableHandle, ref this.response, ref this.rawData);
            seekRowBookmarkResponse = (RopSeekRowBookmarkResponse)this.response;

            bool rowNoLongerVisible = Convert.ToBoolean(seekRowBookmarkResponse.RowNoLongerVisible);
            bool hasSoughtLess = false;
            bool isActualRowCount = false;

            if (wantRowMovedCountout)
            {
                hasSoughtLess = Convert.ToBoolean(seekRowBookmarkResponse.HasSoughtLess);
                if (forwardSeek)
                {
                    if (enoughRow)
                    {
                        isActualRowCount = (uint)(this.rowCount - this.userDefinedBookmarkPosition) == System.Math.Abs(seekRowBookmarkResponse.RowsSought);
                    }
                    else
                    {
                        isActualRowCount = this.rowCount == seekRowBookmarkResponse.RowsSought;
                    }
                }
                else
                {
                    isActualRowCount = (uint)this.userDefinedBookmarkPosition == System.Math.Abs(seekRowBookmarkResponse.RowsSought);
                }
            }
            else
            {
                hasSoughtLess = true;
                isActualRowCount = true;
            }

            if (this.RopSeekRowBookmarkResponse != null)
            {
                this.RopSeekRowBookmarkResponse(wantRowMovedCountout, enoughRow, hasSoughtLess, isActualRowCount, rowNoLongerVisible);
            }

            return (TableRopReturnValues)seekRowBookmarkResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to seek row fractional for a table
        /// </summary>
        /// <param name="zeroNumerator">Identify whether the Numerator is zero</param>
        /// <param name="smallDenominator">Identify whether the Numerator bigger than Denominator</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopSeekRowFractional(bool zeroNumerator, bool smallDenominator)
        {
            RopSeekRowFractionalRequest seekRowFractionalRequest;
            RopSeekRowFractionalResponse seekRowFractionalResponse;

            seekRowFractionalRequest.RopId = 0x1A;
            seekRowFractionalRequest.LogonId = 0x00;
            seekRowFractionalRequest.InputHandleIndex = 0x00;
            if (zeroNumerator)
            {
                seekRowFractionalRequest.Numerator = 0x00000000; // The cursor is set to the first row in the table.
            }
            else
            {
                if (smallDenominator)
                {
                    seekRowFractionalRequest.Numerator = 0x00000003; // Set a value larger than the denominator.
                }
                else
                {
                    seekRowFractionalRequest.Numerator = 0x00000001; // Set a value smaller than the denominator.
                }
            }

            seekRowFractionalRequest.Denominator = 0x00000002;

            this.DoSingleCallROP(seekRowFractionalRequest, this.tableHandle, ref this.response, ref this.rawData);
            seekRowFractionalResponse = (RopSeekRowFractionalResponse)this.response;

            CursorPosition cursorPosition;
            uint currentRowsIndex = this.GetRowsCount(false); // Gets current cursor position.
            if (currentRowsIndex == 0)
            {
                cursorPosition = CursorPosition.BEGIN;
            }
            else if (currentRowsIndex == this.rowCount)
            {
                cursorPosition = CursorPosition.END;
            }
            else
            {
                cursorPosition = CursorPosition.CURRENT;
            }

            if (this.RopSeekRowFractionalResponse != null)
            {
                this.RopSeekRowFractionalResponse(zeroNumerator, smallDenominator, cursorPosition);
            }

            return (TableRopReturnValues)seekRowFractionalResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to create bookmark for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopCreateBookmark()
        {
            this.userDefinedBookmarkPosition = this.GetRowsCount(false);

            RopCreateBookmarkRequest createBookmarkRequest;
            RopCreateBookmarkResponse createBookmarkResponse;

            createBookmarkRequest.RopId = 0x1B;
            createBookmarkRequest.LogonId = 0x00;
            createBookmarkRequest.InputHandleIndex = 0x00;

            this.DoSingleCallROP(createBookmarkRequest, this.tableHandle, ref this.response, ref this.rawData);
            createBookmarkResponse = (RopCreateBookmarkResponse)this.response;

            this.ropType = TableRopType.CREATEBOOKMARK;
            this.userDefinedBookmark = createBookmarkResponse.Bookmark;
            this.userDefinedBookmarkSize = createBookmarkResponse.BookmarkSize;

            return (TableRopReturnValues)createBookmarkResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to query all columns for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopQueryColumnsAll()
        {
            RopQueryColumnsAllRequest queryColumnsAllRequest;
            RopQueryColumnsAllResponse queryColumnsAllResponse;

            queryColumnsAllRequest.RopId = 0x37;
            queryColumnsAllRequest.LogonId = 0x00;
            queryColumnsAllRequest.InputHandleIndex = 0x00;

            this.DoSingleCallROP(queryColumnsAllRequest, this.tableHandle, ref this.response, ref this.rawData);
            queryColumnsAllResponse = (RopQueryColumnsAllResponse)this.response;

            return (TableRopReturnValues)queryColumnsAllResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to find row for a table
        /// </summary>
        /// <param name="findRowFlags">Identify the flags for this operation</param>
        /// <param name="origin">Identify the bookmark</param>
        /// <param name="isValidRestriction">Identify whether the restriction is valid</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopFindRow(FindRowFlags findRowFlags, BookmarkType origin, bool isValidRestriction)
        {
            RopFindRowRequest findRowRequest;
            RopFindRowResponse findRowResponse;

            findRowRequest.RopId = 0x4F;
            findRowRequest.LogonId = 0x00;
            findRowRequest.InputHandleIndex = 0x00;
            findRowRequest.FindRowFlags = (byte)findRowFlags;
            byte[] data;
            if (isValidRestriction)
            {
                data = AdapterHelper.GenerateValidRestrictData(this.Site, this.tableType);
            }
            else
            {
                data = new byte[] { 0x03, 0x02, 0x00, 0x04, 0x00, 0x1f, 0x00, 0x1a, 0x0c, 0x1f, 0x00, 0x1a, 0x0c, 0x24 }; // SenderName is $, which is not exist.
            }

            findRowRequest.RestrictionDataSize = (ushort)data.Length;
            findRowRequest.RestrictionData = data;
            findRowRequest.Origin = (byte)origin;
            if (origin == BookmarkType.BOOKMARK_CUSTOM)
            {
                findRowRequest.Bookmark = this.userDefinedBookmark; // Returned by a previous RopCreateBookmark request.
                findRowRequest.BookmarkSize = this.userDefinedBookmarkSize;
            }
            else
            {
                findRowRequest.Bookmark = null;
                findRowRequest.BookmarkSize = 0x0;
            }

            if (Common.IsRequirementEnabled(827, this.Site))
            {
                this.DoSingleCallROP(findRowRequest, this.tableHandle, ref this.response, ref this.rawData);
            }
            else
            {
                try
                {
                    this.DoSingleCallROP(findRowRequest, this.tableHandle, ref this.response, ref this.rawData);
                }
                catch (Exception ex)
                {
                    Site.Log.Add(LogEntryKind.Comment, ex.Message);
                    Site.Log.Add(LogEntryKind.Comment, "The behavior when a RopFindRow ROP is sent before a successful RopSetColumns ROP is undefined for implementation following Exchange 2007.");
                    return TableRopReturnValues.unexpected;
                }
            }

            findRowResponse = (RopFindRowResponse)this.response;
            bool rowNoLongerVisible = findRowResponse.RowNoLongerVisible == 0x01 ? true : false;

            bool hasRowData = Convert.ToBoolean(findRowResponse.HasRowData);

            bool isCorrectRowData = true;
            bool isLastSetColumnData = true;
            if ((TableRopReturnValues)findRowResponse.ReturnValue == TableRopReturnValues.success)
            {
                if (hasRowData)
                {
                    if (this.tableType == TableType.CONTENT_TABLE)
                    {
                        // Gets the SenderName.
                        string strSenderName = Encoding.Unicode.GetString(findRowResponse.RowData.PropertyValues[2].Value);

                        // TRUE:the senderName is Test1.
                        isCorrectRowData = this.IsExpectedSender(strSenderName, Common.GetConfigurationPropertyValue("Sender1Name", this.Site));
                    }
                    else if (this.tableType == TableType.HIERARCHY_TABLE)
                    {
                        // Gets the DisplayName.
                        string strDisplayName = Encoding.Unicode.GetString(findRowResponse.RowData.PropertyValues[5].Value);

                        // TRUE:the displayName is Test1.
                        isCorrectRowData = strDisplayName == (Constants.TestGetHierarchyTableFolderName1 + '\0');
                    }
                    else if (this.tableType == TableType.RULES_TABLE)
                    {
                        // Gets the RuleName property.
                        string strRuleName = Encoding.Unicode.GetString(findRowResponse.RowData.PropertyValues[1].Value);

                        // TRUE:the RuleName is MarkAsRead.
                        isCorrectRowData = strRuleName == (Constants.RuleNameMarkAsRead + '\0');
                    }

                    isLastSetColumnData = this.latestPropertyTags.Length == findRowResponse.RowData.PropertyValues.Count;
                }
            }

            if (this.RopFindRowResponse != null)
            {
                this.RopFindRowResponse(findRowFlags, isValidRestriction, hasRowData, isCorrectRowData, rowNoLongerVisible, isLastSetColumnData);
            }

            return (TableRopReturnValues)findRowResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to free bookmark for a table.
        /// </summary>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopFreeBookmark()
        {
            RopFreeBookmarkRequest freeBookmarkRequest;
            RopFreeBookmarkResponse freeBookmarkResponse;

            freeBookmarkRequest.RopId = 0x89;
            freeBookmarkRequest.LogonId = 0x00;
            freeBookmarkRequest.InputHandleIndex = 0x00;
            freeBookmarkRequest.BookmarkSize = this.userDefinedBookmark == null ? (ushort)0 : (ushort)this.userDefinedBookmark.Length;
            freeBookmarkRequest.Bookmark = this.userDefinedBookmark; // Returned by a previous RopCreateBookmark request.

            if (this.tableType != TableType.CONTENT_TABLE || Common.IsRequirementEnabled(549, this.Site))
            {
                this.DoSingleCallROP(freeBookmarkRequest, this.tableHandle, ref this.response, ref this.rawData);
                freeBookmarkResponse = (RopFreeBookmarkResponse)this.response;
            }
            else
            {
                Site.Log.Add(LogEntryKind.Comment, "The behavior when RopFreeBookmark is called on a content table object is undefined for implementation following Exchange 2007.");
                return TableRopReturnValues.unexpected;
            }

            return (TableRopReturnValues)freeBookmarkResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to reset a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopResetTable()
        {
            RopResetTableRequest resetTableRequest;
            RopResetTableResponse resetTableResponse;

            resetTableRequest.RopId = 0x81;
            resetTableRequest.LogonId = 0x00;
            resetTableRequest.InputHandleIndex = 0x00;

            this.DoSingleCallROP(resetTableRequest, this.tableHandle, ref this.response, ref this.rawData);
            resetTableResponse = (RopResetTableResponse)this.response;

            bool isPreviousDataset = true;

            if (this.ropType == TableRopType.MOVECURSOR || this.ropType == TableRopType.RESTRICT)
            {
                isPreviousDataset = this.GetRowsCount(false) == 0;
            }

            if (this.RopResetTableResponse != null)
            {
                this.RopResetTableResponse(this.ropType, isPreviousDataset);
            }

            return (TableRopReturnValues)resetTableResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to expand row for a table
        /// </summary>
        /// <param name="nonexistCategoryId">Identify whether the CategoryId row exists</param>
        /// <param name="noncollapsedCategoryId">Identify whether the CategoryId row is not collapsed</param>
        /// <param name="isMaxRowCountInvalid">Identify whether the MaxRowCount is invalid</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopExpandRow(bool nonexistCategoryId, bool noncollapsedCategoryId, bool isMaxRowCountInvalid)
        {
            RopExpandRowRequest expandRowRequest;
            RopExpandRowResponse expandRowResponse;

            byte[] propertyValue = new byte[] { 0x01, 0x00, 0x00, 0x00, 0x13, 0x98, 0x24, 0x13 }; // Initial an undefined value.
            ulong index = 1;
            ulong categoryId = 0;
            if (this.tableType == TableType.CONTENT_TABLE)
            {
                propertyValue = this.propertyValues[0].Value; // PidTagInstID.
                if (noncollapsedCategoryId)
                {
                    if (!this.isExpanded)
                    {
                        foreach (byte a in propertyValue)
                        {
                            categoryId += (ulong)a * index;
                            index *= 0x100;
                        }

                        expandRowRequest.RopId = 0x59;
                        expandRowRequest.LogonId = 0x00;
                        expandRowRequest.InputHandleIndex = 0x00;
                        expandRowRequest.MaxRowCount = this.maxRowCountInExpandRowRequest; // The maximum number of leaf rows to expand.
                        expandRowRequest.CategoryId = categoryId;

                        this.DoSingleCallROP(expandRowRequest, this.tableHandle, ref this.response, ref this.rawData);
                        expandRowResponse = (RopExpandRowResponse)this.response;
                    }
                }
                else
                {
                    if (this.isExpanded)
                    {
                        RopCollapseRowRequest collapseRowRequest;

                        index = 1;
                        categoryId = 0;
                        foreach (byte a in propertyValue)
                        {
                            categoryId += (ulong)a * index;
                            index *= 0x100;
                        }

                        collapseRowRequest.RopId = 0x5A;
                        collapseRowRequest.LogonId = 0x00;
                        collapseRowRequest.InputHandleIndex = 0x00;
                        collapseRowRequest.CategoryId = categoryId;

                        this.DoSingleCallROP(collapseRowRequest, this.tableHandle, ref this.response, ref this.rawData);
                    }
                }
            }

            if (nonexistCategoryId)
            {
                propertyValue = new byte[] { 0x00 }; // Not exist categoryId.
            }

            index = 1;
            categoryId = 0;
            foreach (byte a in propertyValue)
            {
                categoryId += (ulong)a * index;
                index *= 0x100;
            }

            expandRowRequest.RopId = 0x59;
            expandRowRequest.LogonId = 0x00;
            expandRowRequest.InputHandleIndex = 0x00;

            if (isMaxRowCountInvalid)
            {
                expandRowRequest.MaxRowCount = 1; // The maximum is set to invalid.
            }
            else
            {
                expandRowRequest.MaxRowCount = this.maxRowCountInExpandRowRequest; // The maximum number of leaf rows to expand.
            }

            expandRowRequest.CategoryId = categoryId;

            this.DoSingleCallROP(expandRowRequest, this.tableHandle, ref this.response, ref this.rawData);
            expandRowResponse = (RopExpandRowResponse)this.response;
            bool isLastSetColumnData = false;
            if ((TableRopReturnValues)expandRowResponse.ReturnValue == TableRopReturnValues.success)
            {
                this.isExpanded = true;
                if (expandRowResponse.RowData.PropertyRows != null)
                {
                    isLastSetColumnData = this.latestPropertyTags.Length == expandRowResponse.RowData.PropertyRows[0].PropertyValues.Count;
                }
            }

            if (this.RopExpandRowResponse != null)
            {
                this.RopExpandRowResponse(isLastSetColumnData);
            }

            return (TableRopReturnValues)expandRowResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to Collapse row for a table
        /// </summary>
        /// <param name="nonexistCategoryId">Identify whether the CategoryId row exists</param>
        /// <param name="nonexpandedCategoryId">Identify whether the CategoryId row is not expanded</param>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopCollapseRow(bool nonexistCategoryId, bool nonexpandedCategoryId)
        {
            RopCollapseRowRequest collapseRowRequest;
            RopCollapseRowResponse collapseRowResponse;
            byte[] propertyValue = new byte[] { 0x01, 0x00, 0x00, 0x00, 0x13, 0x98, 0x24, 0x13 }; // Initial an undefined value.
            ulong index = 1;
            ulong categoryId = 0;
            if (this.tableType == TableType.CONTENT_TABLE)
            {
                propertyValue = this.propertyValues[0].Value; // PidTagInstID.
                if (nonexpandedCategoryId)
                {
                    if (this.isExpanded)
                    {
                        index = 1;
                        categoryId = 0;
                        foreach (byte a in propertyValue)
                        {
                            categoryId += (ulong)a * index;
                            index *= 0x100;
                        }

                        collapseRowRequest.RopId = 0x5A;
                        collapseRowRequest.LogonId = 0x00;
                        collapseRowRequest.InputHandleIndex = 0x00;
                        collapseRowRequest.CategoryId = categoryId;

                        this.DoSingleCallROP(collapseRowRequest, this.tableHandle, ref this.response, ref this.rawData);
                        collapseRowResponse = (RopCollapseRowResponse)this.response;
                    }
                }
                else
                {
                    if (!this.isExpanded)
                    {
                        RopExpandRowRequest expandRowRequest;

                        index = 1;
                        categoryId = 0;
                        foreach (byte a in propertyValue)
                        {
                            categoryId += (ulong)a * index;
                            index *= 0x100;
                        }

                        expandRowRequest.RopId = 0x59;
                        expandRowRequest.LogonId = 0x00;
                        expandRowRequest.InputHandleIndex = 0x00;
                        expandRowRequest.MaxRowCount = this.maxRowCountInExpandRowRequest; // The maximum number of leaf rows to expand.
                        expandRowRequest.CategoryId = categoryId;

                        this.DoSingleCallROP(expandRowRequest, this.tableHandle, ref this.response, ref this.rawData);
                    }
                }
            }

            if (nonexistCategoryId)
            {
                propertyValue = new byte[] { 0x00 }; // Not exist categoryId.
            }

            index = 1;
            categoryId = 0;
            foreach (byte a in propertyValue)
            {
                categoryId += (ulong)a * index;
                index *= 0x100;
            }

            collapseRowRequest.RopId = 0x5A;
            collapseRowRequest.LogonId = 0x00;
            collapseRowRequest.InputHandleIndex = 0x00;
            collapseRowRequest.CategoryId = categoryId;

            this.DoSingleCallROP(collapseRowRequest, this.tableHandle, ref this.response, ref this.rawData);
            collapseRowResponse = (RopCollapseRowResponse)this.response;

            if ((TableRopReturnValues)collapseRowResponse.ReturnValue == TableRopReturnValues.success)
            {
                this.isExpanded = false;
            }

            return (TableRopReturnValues)collapseRowResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to get collapse state for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopGetCollapseState()
        {
            RopGetCollapseStateRequest getCollapseStateRequest;
            RopGetCollapseStateResponse getCollapseStateResponse;

            byte[] propertyValue = new byte[] { 0x01, 0x00, 0x00, 0x00, 0x13, 0x98, 0x24, 0x13 }; // Initial an undefined value.
            if (this.tableType == TableType.CONTENT_TABLE)
            {
                propertyValue = this.propertyValues[0].Value; // PidTagInstID.
            }

            ulong index = 1;
            ulong categoryId = 0;
            foreach (byte a in propertyValue)
            {
                categoryId += (ulong)a * index;
                index *= 0x100;
            }

            getCollapseStateRequest.RopId = 0x6B;
            getCollapseStateRequest.LogonId = 0x00;
            getCollapseStateRequest.InputHandleIndex = 0x00;
            getCollapseStateRequest.RowId = categoryId;
            getCollapseStateRequest.RowInstanceNumber = 0x00000000;

            this.DoSingleCallROP(getCollapseStateRequest, this.tableHandle, ref this.response, ref this.rawData);
            getCollapseStateResponse = (RopGetCollapseStateResponse)this.response;

            if (!this.isCollapseStateGot)
            {
                this.collapseState = getCollapseStateResponse.CollapseState;
                this.collapseStateSize = getCollapseStateResponse.CollapseStateSize;
                this.isCollapseStateGot = true;
            }

            bool isCollapseStateRebuildCorrectly = false;
            if (this.isSetCollapseStateCalled)
            {
                if (this.collapseStateSize == getCollapseStateResponse.CollapseStateSize)
                {
                    int i = 0;
                    for (; i < this.collapseStateSize; i++)
                    {
                        if (this.collapseState[i] == getCollapseStateResponse.CollapseState[i])
                        {
                            continue;
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (i == this.collapseStateSize)
                    {
                        isCollapseStateRebuildCorrectly = true;
                    }
                }
                else
                {
                    isCollapseStateRebuildCorrectly = false;
                }
            }

            if (this.RopGetCollapseStateResponse != null)
            {
                this.RopGetCollapseStateResponse(this.isSetCollapseStateCalled, isCollapseStateRebuildCorrectly);
            }

            return (TableRopReturnValues)getCollapseStateResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to set collapse state for a table
        /// </summary>
        /// <returns>Table ROP return value</returns>
        public TableRopReturnValues RopSetCollapseState()
        {
            RopSetCollapseStateRequest setCollapseStateRequest;
            RopSetCollapseStateResponse setCollapseStateResponse;

            setCollapseStateRequest.RopId = 0x6C;
            setCollapseStateRequest.LogonId = 0x00;
            setCollapseStateRequest.InputHandleIndex = 0x00;
            setCollapseStateRequest.CollapseStateSize = this.collapseStateSize;
            setCollapseStateRequest.CollapseState = this.collapseState; // This data is obtained by sending a RopGetCollapseState request.

            this.DoSingleCallROP(setCollapseStateRequest, this.tableHandle, ref this.response, ref this.rawData);
            setCollapseStateResponse = (RopSetCollapseStateResponse)this.response;
            if ((TableRopReturnValues)setCollapseStateResponse.ReturnValue == TableRopReturnValues.success)
            {
                this.isSetCollapseStateCalled = true;
                this.userDefinedBookmarkSize = setCollapseStateResponse.BookmarkSize;
                this.userDefinedBookmark = setCollapseStateResponse.Bookmark;
            }

            return (TableRopReturnValues)setCollapseStateResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to release a table
        /// </summary>
        public void RopRelease()
        {
            RopReleaseRequest ropReleaseRequest;

            ropReleaseRequest.RopId = 0x01;
            ropReleaseRequest.LogonId = 0x00;
            ropReleaseRequest.InputHandleIndex = 0x00;

            this.DoSingleCallROP(ropReleaseRequest, this.tableHandle, ref this.response, ref this.rawData);
        }
        #endregion

        /// <summary>
        /// Determines if the requirement is enabled or not.
        /// </summary>
        /// <param name="requirementId">The requirement Id</param>
        /// <param name="enabled">Requirement is enable or not.</param>
        public void CheckRequirementEnabled(int requirementId, out bool enabled)
        {
            enabled = Common.IsRequirementEnabled(requirementId, this.Site);

            // If the requirement number is 910 with false value, which means the server doesn't support RopSeekRowFractional ROP.
            if (requirementId == 910 && enabled == false)
            {
                this.needDoCleanup = false;

                // Exchange server 2007 does not support the RopSeekRowFractional ROP.
                Site.Assume.Inconclusive("The Exchange server 2007 doesn't support the RopSeekRowFractional ROP.");
            }
        }

        /// <summary>
        /// This method is used to check whether MAPIHTTP transport is supported by SUT.
        /// </summary>
        /// <param name="isSupported">The transport is supported or not.</param>
        public void CheckMAPIHTTPTransportSupported(out bool isSupported)
        {
            // Check whether support MS-OXCMAPIHTTP transport.
            if ((Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http") && (!Common.IsRequirementEnabled(1340, this.Site)))
            {
                this.needDoCleanup = false;
                isSupported = false;
                Site.Assert.Inconclusive("This case can not run with MAPIHTTP transport with current SUT version. MAPIHTTP is supported from Exchange Server 2013 SP1.");
            }
            else
            {
                isSupported = true;
            }
        }

        #endregion

        #region MS_OXCTABLAdapter members
        
        /// <summary>
        /// Initialize the adapter.
        /// </summary>
        /// <param name="testSite">Test site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            AdapterHelper.Site = testSite;
            testSite.DefaultProtocolDocShortName = "MS-OXCTABL";
            Common.MergeConfiguration(this.Site);
            this.maxRowCountInExpandRowRequest = 0x32;

            // Implementation following Exchange 2013 does not support a value greater than 0 for the MaxRowCount field.
            if (Common.IsRequirementEnabled(748, this.Site))
            {
                this.maxRowCountInExpandRowRequest = 0x00;
            }

            this.oxcropsClient = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));

            this.waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
        }

        /// <summary>
        /// Reset the adapter.
        /// </summary>
        public override void Reset()
        {
            if (this.needDoCleanup)
            {
                if (this.tableType == TableType.CONTENT_TABLE || this.tableType == TableType.RULES_TABLE || this.tableType == TableType.HIERARCHY_TABLE || this.tableType == TableType.ATTACHMENTS_TABLE)
                {
                    this.CleanInbox();
                }

                if (this.isConnected)
                {
                    this.Disconnect();
                }
            }

            this.latestSortOrder = SortOrderFlag.NotSort;
            this.latestSuccessSortOrder = SortOrderFlag.NotSort;
            this.latestRestrict = RestrictFlag.NotRestrict;
            this.latestSuccessRestrict = RestrictFlag.NotRestrict;
            this.areMultipleSortOrders = false;
            this.areAllSortOrdersUsedAsCategory = false;
            this.areAllCategoryExpanded = false;
            this.isExpanded = false;
            this.isSetCollapseStateCalled = false;
            this.isCollapseStateGot = false;
            this.globalIsRestrictAsynchronous = false;
            this.globalIsSetColumnsAsynchronous = false;
            this.globalIsSortTableAsynchronous = false;

            this.userDefinedBookmark = new byte[] { 0x00 };
            this.userDefinedBookmarkPosition = 0;
            this.userDefinedBookmarkSize = 1;

            this.propertyValues = new List<PropertyValue>();
            this.collapseState = new byte[] { 0x00 };
            this.collapseStateSize = 1;
            this.needDoCleanup = true;

            base.Reset();
        }

        #endregion

        #region Initialize the table
        /// <summary>
        /// This method is used to initialize the table
        /// </summary>
        /// <param name="tableType">Identify the table type</param>
        public void InitializeTable(TableType tableType)
        {
            this.tableType = tableType; // Record the type of table.

            // Connect to the server.
            string serverName = string.Empty;
            serverName = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
            this.Connect(serverName);

            // Log on to the server.
            RopLogonResponse logonResponse = this.Logon(out this.inputObjHandle);

            // Initial the folder to Inbox.
            this.folderID = logonResponse.FolderIds[4];

            // Initial the sign to true.
            this.sign = true;

            // Prepare the table variable
            this.PrepareTable(this.tableType);

            this.userDefinedBookmark = new byte[] { 0x00 };
            this.userDefinedBookmarkPosition = 0;
            this.userDefinedBookmarkSize = 1;

            this.collapseState = new byte[] { 0x00 };
            this.collapseStateSize = 1;

            this.latestSortOrder = SortOrderFlag.NotSort;
            this.latestSuccessSortOrder = SortOrderFlag.NotSort;
            this.latestRestrict = RestrictFlag.NotRestrict;
            this.latestSuccessRestrict = RestrictFlag.NotRestrict;
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Create Sample ContentsTable PropertyTags
        /// </summary>
        /// <returns>Return PropertyTag array</returns>
        private PropertyTag[] CreateSampleContentsTablePropertyTags()
        {
            PropertyTag[] propertyTags = new PropertyTag[10];
            PropertyTag tag;

            // PidTagInstID
            tag.PropertyId = 0x674D;
            tag.PropertyType = 0x0014;
            propertyTags[0] = tag;

            // PidTagInstanceNum
            tag.PropertyId = 0x674e;
            tag.PropertyType = 0x0003;
            propertyTags[1] = tag;

            // PidTagSenderName
            tag.PropertyId = 0x0C1A;
            tag.PropertyType = 0x001F;
            propertyTags[2] = tag;

            // PidTagRowType
            tag.PropertyId = 0x0FF5;
            tag.PropertyType = 0x0003;
            propertyTags[3] = tag;

            // PidTagDepth
            tag.PropertyId = 0x3005;
            tag.PropertyType = 0x0003;
            propertyTags[4] = tag;

            // PidTagContentCount
            tag.PropertyId = 0x3602;
            tag.PropertyType = 0x0003;
            propertyTags[5] = tag;

            // PidTagContentUnreadCount
            tag.PropertyId = 0x3602;
            tag.PropertyType = 0x0003;
            propertyTags[6] = tag;

            // PidTagSubject
            tag.PropertyId = 0x0037;
            tag.PropertyType = 0x001F;
            propertyTags[7] = tag;

            // PidTagMessageDeliveryTime
            tag.PropertyId = 0x0e06;
            tag.PropertyType = 0x0040;
            propertyTags[8] = tag;

            // PidTagMessageSize
            tag.PropertyId = 0x0E08;
            tag.PropertyType = 0x0003;
            propertyTags[9] = tag;

            return propertyTags;
        }

        /// <summary>
        /// Create another sample ContentsTable PropertyTags which have no error value
        /// </summary>
        /// <returns>Return PropertyTag array</returns>
        private PropertyTag[] CreateSampleContentsTablePropertyTagsWithNoError()
        {
            PropertyTag[] propertyTags = new PropertyTag[6];
            PropertyTag tag;

            // PidTagInstID
            tag.PropertyId = 0x674d;
            tag.PropertyType = 0x0014;
            propertyTags[0] = tag;

            // PidTagInstanceNum
            tag.PropertyId = 0x674e;
            tag.PropertyType = 0x0003;
            propertyTags[1] = tag;

            // PidTagSenderName
            tag.PropertyId = 0x0C1A;
            tag.PropertyType = 0x001F;
            propertyTags[2] = tag;

            // PidTagRowType
            tag.PropertyId = 0x0FF5;
            tag.PropertyType = 0x0003;
            propertyTags[3] = tag;

            // PidTagDepth
            tag.PropertyId = 0x3005;
            tag.PropertyType = 0x0003;
            propertyTags[4] = tag;

            // PidTagMessageDeliveryTime
            tag.PropertyId = 0x0e06;
            tag.PropertyType = 0x0040;
            propertyTags[5] = tag;

            return propertyTags;
        }

        /// <summary>
        /// Create Sample Hierarchy PropertyTags
        /// </summary>
        /// <returns>Return PropertyTag array</returns>
        private PropertyTag[] CreateSampleHierarchyTablePropertyTags()
        {
            PropertyTag[] propertyTags = new PropertyTag[7];
            PropertyTag tag;

            // PidTagAddressBookEntryId
            tag.PropertyId = 0x663b;
            tag.PropertyType = 0x0102;
            propertyTags[0] = tag;

            // PidTagFolderId
            tag.PropertyId = 0x6748;
            tag.PropertyType = 0x0014;
            propertyTags[1] = tag;

            // PidTagParentEntryId
            tag.PropertyId = 0x0e09;
            tag.PropertyType = 0x0102;
            propertyTags[2] = tag;

            // PidTagHierarchyChangeNumber
            tag.PropertyId = 0x663e;
            tag.PropertyType = 0x0003;
            propertyTags[3] = tag;

            // PidTagSubfolders
            tag.PropertyId = 0x360a;
            tag.PropertyType = 0x000b;
            propertyTags[4] = tag;

            // PidTagDisplayName
            tag.PropertyId = 0x3001;
            tag.PropertyType = 0x001f;
            propertyTags[5] = tag;

            // PidTagFolderType
            tag.PropertyId = 0x3601;
            tag.PropertyType = 0x0003;
            propertyTags[6] = tag;

            return propertyTags;
        }

        /// <summary>
        /// Create Rules Table PropertyTags
        /// </summary>
        /// <returns>Return Invalid PropertyTag array</returns>
        private PropertyTag[] CreateRulesTablePropertyTags()
        {
            PropertyTag[] propertyTags = new PropertyTag[3];
            PropertyTag tag;

            // PropertyId of PidTagRuleActions
            tag.PropertyId = 0x6680;
            tag.PropertyType = 0x00FE;
            propertyTags[0] = tag;

            // PropertyId of PidTagRuleName
            tag.PropertyId = 0x6682;
            tag.PropertyType = 0x001F;
            propertyTags[1] = tag;

            // PropertyId of PidTagRuleProvider
            tag.PropertyId = 0x6681;
            tag.PropertyType = 0x001F;
            propertyTags[2] = tag;

            return propertyTags;
        }

        /// <summary>
        /// Create invalid contentsTable PropertyTags
        /// </summary>
        /// <returns>Return Invalid PropertyTag array</returns>
        private PropertyTag[] CreateInvalidContentsTablePropertyTags()
        {
            PropertyTag[] propertyTags = new PropertyTag[1];
            PropertyTag tag;

            // The propertyTag is Invalid, specified in [MS-OXCTABL] section 3.2.5.2
            tag.PropertyId = 0x8007;
            tag.PropertyType = 0x0057;
            propertyTags[0] = tag;

            return propertyTags;
        }

        /// <summary>
        /// Create Sample SortOrders in ascending order
        /// </summary>
        /// <param name="multipleSortOrders">Indicate whether sort orders contains more than one sort order structure</param>
        /// <param name="isMaximumCategory">Indicate whether sort orders contains a sort order structure with the order set to MaximumCategory</param>
        /// <returns>Return SortOrder array</returns>
        private SortOrder[] CreateSampleSortOrdersAscending(bool multipleSortOrders, bool isMaximumCategory)
        {
            SortOrder[] sortOrders;
            if (!multipleSortOrders)
            {
                sortOrders = new SortOrder[1];
            }
            else
            {
                sortOrders = new SortOrder[2];
            }

            SortOrder sortOrder;

            // PidTagSenderName
            sortOrder.PropertyId = 0x0C1A;
            sortOrder.PropertyType = 0x001F;
            sortOrder.Order = 0x00; // Ascending
            sortOrders[0] = sortOrder;

            if (multipleSortOrders)
            {
                // PidTagMessageDeliveryTime
                sortOrder.PropertyId = 0x0e06;
                sortOrder.PropertyType = 0x0040;
                if (isMaximumCategory)
                {
                    sortOrder.Order = 0x04; // MaximumCategory
                }
                else
                {
                    sortOrder.Order = 0x00; // Ascending
                }

                sortOrders[1] = sortOrder;
            }

            return sortOrders;
        }

        /// <summary>
        /// Create Sample SortOrders in descending order
        /// </summary>
        /// <param name="multipleSortOrders">Indicate whether sort orders contains more than one sort order structure</param>
        /// <returns>Return SortOrder array</returns>
        private SortOrder[] CreateSampleSortOrdersDescending(bool multipleSortOrders)
        {
            SortOrder[] sortOrders;
            if (!multipleSortOrders)
            {
                sortOrders = new SortOrder[1];
            }
            else
            {
                sortOrders = new SortOrder[2];
            }

            SortOrder sortOrder;

            // PidTagSenderName
            sortOrder.PropertyId = 0x0C1A;
            sortOrder.PropertyType = (ushort)0x001F;
            sortOrder.Order = 0x01; // Descending
            sortOrders[0] = sortOrder;

            if (multipleSortOrders)
            {
                // PidTagMessageDeliveryTime
                sortOrder.PropertyId = 0x0e06;
                sortOrder.PropertyType = 0x0040;
                sortOrder.Order = 0x01; // Descending
                sortOrders[1] = sortOrder;
            }

            return sortOrders;
        }

        /// <summary>
        /// Gets the number of rows from current location of cursor
        /// </summary>
        /// <param name="forwardRead">Specifies the direction to read rows</param>
        /// <returns>Returns the number of rows</returns>
        private uint GetRowsCount(bool forwardRead)
        {
            RopQueryPositionRequest queryPositionRequest;
            RopQueryPositionResponse queryPositionResponse;

            queryPositionRequest.RopId = 0x17;
            queryPositionRequest.LogonId = 0x00;
            queryPositionRequest.InputHandleIndex = 0x00;

            this.responseSOHs = this.DoSingleCallROP(queryPositionRequest, this.tableHandle, ref this.response, ref this.rawData);
            queryPositionResponse = (RopQueryPositionResponse)this.response;

            uint numerator = 0;
            uint denominator = 0;
            if ((TableRopReturnValues)queryPositionResponse.ReturnValue == TableRopReturnValues.success)
            {
                numerator = queryPositionResponse.Numerator;
                denominator = queryPositionResponse.Denominator;
                this.rowCount = denominator;
                if (forwardRead)
                {
                    return denominator - numerator;
                }
                else
                {
                    return numerator;
                }
            }

            return 0;
        }

        /// <summary>
        /// Connect to the server for RPC calling or HTTP request.
        /// </summary>
        /// <param name="serverName">Server to connect.</param>
        /// <returns>Result of connecting.</returns>
        private bool Connect(string serverName)
        {
            this.isConnected = this.oxcropsClient.Connect(
                    serverName,
                    ConnectionType.PrivateMailboxServer,
                    Common.GetConfigurationPropertyValue("Sender1Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("Sender1Name", this.Site),
                    Common.GetConfigurationPropertyValue("Sender1Password", this.Site));

            return this.isConnected;
        }

        /// <summary>
        /// Disconnect from the server.
        /// </summary>
        public void Disconnect()
        {
            if (this.tableType == TableType.CONTENT_TABLE || this.tableType == TableType.RULES_TABLE || this.tableType == TableType.HIERARCHY_TABLE || this.tableType == TableType.ATTACHMENTS_TABLE)
            {
                this.CleanInbox();
            }

            bool ret = this.oxcropsClient.Disconnect();
            if (ret)
            {
                this.isConnected = false;
            }
        }

        /// <summary>
        /// Send ROP request with single operation.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inputObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <returns>Server objects handles in response.</returns>
        private List<List<uint>> DoSingleCallROP(ISerializable ropRequest, uint inputObjHandle, ref object response, ref byte[] rawData)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };

            List<uint> requestSOH = new List<uint>
            {
                inputObjHandle
            };

            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle, set default value to 0xFFFFFFFF.
                requestSOH.Add(0xFFFFFFFF);
            }

            List<IDeserializable> responseRops = new List<IDeserializable>();
            List<List<uint>> responseSOHs = new List<List<uint>>();

            uint ret = this.oxcropsClient.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, 0x10008);
            if (ret != OxcRpcErrorCode.ECNone)
            {
                return responseSOHs;
            }

            if (responseRops != null)
            {
                if (responseRops.Count > 0)
                {
                    response = responseRops[0];

                    // Call route response method to parse and verify the response structures.
                    this.ProcessResponse(responseRops[0]);
                }
            }
            else
            {
                response = null;
            }

            this.VerifyMAPITransport();
            return responseSOHs;
        }

        /// <summary>
        /// This method is used to process the response and 
        /// call the corresponding verification method.
        /// </summary>
        /// <param name="response">The IDeserializable type response</param>
        private void ProcessResponse(IDeserializable response)
        {
            if (response == null)
            {
                Site.Assert.Fail("Response should not be null");
            }

            string responseName = response.GetType().Name;

            switch (responseName)
            {
                case "RopAbortResponse":
                    {
                        // Verify RopAbortResponse Response
                        RopAbortResponse abortResponse = (RopAbortResponse)response;
                        this.VerifyRopAbortResponse(abortResponse);
                        break;
                    }

                case "RopCollapseRowResponse":
                    {
                        // Verify RopCollapseRow Response
                        RopCollapseRowResponse collapseRowResponse = (RopCollapseRowResponse)response;
                        this.VerifyRopCollapseRowResponse(collapseRowResponse);
                        break;
                    }

                case "RopCreateBookmarkResponse":
                    {
                        // Verify RopCreateBookmark Response
                        RopCreateBookmarkResponse createBookmarkResponse = (RopCreateBookmarkResponse)response;
                        if (createBookmarkResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopCreateBookmarkResponse(createBookmarkResponse);
                        }

                        break;
                    }

                case "RopExpandRowResponse":
                    {
                        // Verify RopExpandRow Response
                        RopExpandRowResponse expandRowResponse = (RopExpandRowResponse)response;
                        if (expandRowResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopExpandRowResponse(expandRowResponse, this.maxRowCountInExpandRowRequest);
                        }

                        break;
                    }

                case "RopFindRowResponse":
                    {
                        // Verify RopFindRow Response
                        RopFindRowResponse findRowResponse = (RopFindRowResponse)response;

                        this.VerifyRopFindRowResponse(findRowResponse);
                        break;
                    }

                case "RopFreeBookmarkResponse":
                    {
                        break;
                    }

                case "RopGetCollapseStateResponse":
                    {
                        // Verify RopGetCollapseState Response
                        RopGetCollapseStateResponse getCollapseStateResponse = (RopGetCollapseStateResponse)response;
                        if (getCollapseStateResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopGetCollapseStateResponse(getCollapseStateResponse);
                        }

                        break;
                    }

                case "RopGetStatusResponse":
                    {
                        // Verify RopGetStatusResponse Response
                        RopGetStatusResponse getStatusResponse = (RopGetStatusResponse)response;
                        this.VerifyGetStatusResponse(getStatusResponse);
                        break;
                    }

                case "RopQueryColumnsAllResponse":
                    {
                        // Verify RopQueryColumnsALL Response.
                        RopQueryColumnsAllResponse queryColumnsAllResponse = (RopQueryColumnsAllResponse)response;
                        if (queryColumnsAllResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopQueryColumnsALLResponse(queryColumnsAllResponse);
                        }

                        break;
                    }

                case "RopQueryPositionResponse":
                    {
                        // Verify RopQueryPosition Response.
                        RopQueryPositionResponse queryPositionRespone = (RopQueryPositionResponse)response;
                        if (queryPositionRespone.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopQueryPositionResponse(queryPositionRespone);
                        }

                        break;
                    }

                case "RopQueryRowsResponse":
                    {
                        // Verify RopQueryRows Response.
                        RopQueryRowsResponse queryRowsRespone = (RopQueryRowsResponse)response;
                        if (queryRowsRespone.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopQueryRowsResponse(queryRowsRespone, this.rowCountInQueryRowsRequest);
                        }

                        break;
                    }

                case "RopResetTableResponse":
                    {
                        break;
                    }

                case "RopRestrictResponse":
                    {
                        // Verify RopRestrict Response.
                        RopRestrictResponse restrictResponse = (RopRestrictResponse)response;
                        if (restrictResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopRestrictResponse(restrictResponse);
                        }

                        break;
                    }

                case "RopSeekRowResponse":
                    {
                        // Verify RopSeekRow Response.
                        RopSeekRowResponse seekRowResponse = (RopSeekRowResponse)response;
                        if (seekRowResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopSeekRowResponse(seekRowResponse, this.rowCountInSeekRowRequest);
                        }

                        break;
                    }

                case "RopSeekRowBookmarkResponse":
                    {
                        // Verify RopSeekRowBookmark Response.
                        RopSeekRowBookmarkResponse seekRowBookmarkResponse = (RopSeekRowBookmarkResponse)response;
                        if (seekRowBookmarkResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopSeekRowBookmarkResponse(seekRowBookmarkResponse, this.rowCountInSeekRowBookmarkRequest);
                        }

                        break;
                    }

                case "RopSeekRowFractionalResponse":
                    {
                        break;
                    }

                case "RopSetCollapseStateResponse":
                    {
                        // Verify RopSetCollapseState Response.
                        RopSetCollapseStateResponse setCollapsStateResponse = (RopSetCollapseStateResponse)response;
                        if (setCollapsStateResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopSetCollapseStateResponse(setCollapsStateResponse);
                        }

                        break;
                    }

                case "RopSetColumnsResponse":
                    {
                        // Verify RopSetColumns Response.
                        RopSetColumnsResponse setColumnsResponse = (RopSetColumnsResponse)response;
                        if (setColumnsResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopSetColumnsResponse(setColumnsResponse);
                        }

                        break;
                    }

                case "RopSortTableResponse":
                    {
                        // Verify RopSortTable Response.
                        RopSortTableResponse sortTableResponse = (RopSortTableResponse)response;
                        if (sortTableResponse.ReturnValue == 0x00000000)
                        {
                            this.VerifyRopSortTableResponse(sortTableResponse);
                        }

                        break;
                    }

                default:
                    break;
            }
        }

        /// <summary>
        /// This method is used to logon the server
        /// </summary>
        /// <param name="objHandle">The logon object handle</param>
        /// <returns>Rop logon response</returns>
        private RopLogonResponse Logon(out uint objHandle)
        {
            this.inputObjHandle = 0;
            RopLogonRequest logonRequest;

            logonRequest.RopId = 0xFE;
            logonRequest.LogonId = 0x0;
            logonRequest.OutputHandleIndex = 0x0;

            string userDN = Common.GetConfigurationPropertyValue("Sender1Essdn", this.Site) + "\0";

            logonRequest.StoreState = 0;
            logonRequest.LogonFlags = 0x01; // Logon to a private mailbox.
            logonRequest.OpenFlags = 0x01000000;
            logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(userDN);
            logonRequest.Essdn = Encoding.ASCII.GetBytes(userDN);

            this.responseSOHs = this.DoSingleCallROP(logonRequest, this.inputObjHandle, ref this.response, ref this.rawData);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            objHandle = this.responseSOHs[0][logonResponse.OutputHandleIndex];

            return logonResponse;
        }

        /// <summary>
        /// This method is used to get the table handle
        /// </summary>
        /// <param name="tableType">Identify the table type</param>
        private void GetTableHandle(TableType tableType)
        {
            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;
            RopQueryRowsRequest queryRowsRequest;
            RopQueryRowsResponse queryRowsResponse;
            RopSetColumnsRequest setColumnsRequest;
            string server = string.Empty;
            string domain = string.Empty;
            string user1 = string.Empty;
            string user2 = string.Empty;
            PropertyTag[] tempPropertyTags = null;
            switch (tableType)
            {
                case TableType.CONTENT_TABLE:
                    server = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
                    domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
                    user1 = Common.GetConfigurationPropertyValue("Sender1Name", this.Site);
                    user2 = Common.GetConfigurationPropertyValue("Sender2Name", this.Site);
                    for (int i = 0; i < 2; i++)
                    {
                        this.CreateMessageForPrepareTable(user1);
                        this.CreateMessageForPrepareTable(user2);
                    }

                    this.VerifyMailCount(4);
                    RopGetContentsTableRequest getContentsTableRequest;
                    RopGetContentsTableResponse getContentsTableResponse;

                    getContentsTableRequest.RopId = 0x05;
                    getContentsTableRequest.LogonId = 0x00;
                    getContentsTableRequest.InputHandleIndex = 0x00;
                    getContentsTableRequest.OutputHandleIndex = 0x01;
                    getContentsTableRequest.TableFlags = 0x00; // Depth.

                    this.responseSOHs = this.DoSingleCallROP(getContentsTableRequest, this.folderHandle, ref this.response, ref this.rawData);
                    getContentsTableResponse = (RopGetContentsTableResponse)this.response;
                    this.tableHandle = this.responseSOHs[0][getContentsTableResponse.OutputHandleIndex];
                    this.rowCount = getContentsTableResponse.RowCount;

                    break;
                case TableType.ATTACHMENTS_TABLE:
                    server = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
                    domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
                    user1 = Common.GetConfigurationPropertyValue("Sender1Name", this.Site);
                    user2 = Common.GetConfigurationPropertyValue("Sender2Name", this.Site);
                    for (int i = 0; i < 2; i++)
                    {
                        this.CreateMessageForPrepareTable(user1);
                        this.CreateMessageForPrepareTable(user2);
                    }

                    this.VerifyMailCount(4);
                    RopGetContentsTableRequest getContentsTableRequest1;
                    RopGetContentsTableResponse getContentsTableResponse1;

                    getContentsTableRequest1.RopId = 0x05;
                    getContentsTableRequest1.LogonId = 0x00;
                    getContentsTableRequest1.InputHandleIndex = 0x00;
                    getContentsTableRequest1.OutputHandleIndex = 0x01;
                    getContentsTableRequest1.TableFlags = 0x00; // Depth.

                    this.responseSOHs = this.DoSingleCallROP(getContentsTableRequest1, this.folderHandle, ref this.response, ref this.rawData);
                    getContentsTableResponse1 = (RopGetContentsTableResponse)this.response;
                    this.tableHandle = this.responseSOHs[0][getContentsTableResponse1.OutputHandleIndex];

                    tempPropertyTags = this.CreateSampleContentsTablePropertyTags();
                    setColumnsRequest.RopId = 0x12;
                    setColumnsRequest.LogonId = 0x00;
                    setColumnsRequest.InputHandleIndex = 0x00;
                    setColumnsRequest.SetColumnsFlags = 0x00;
                    setColumnsRequest.PropertyTagCount = (ushort)tempPropertyTags.Length;
                    setColumnsRequest.PropertyTags = tempPropertyTags;

                    this.responseSOHs = this.DoSingleCallROP(setColumnsRequest, this.tableHandle, ref this.response, ref this.rawData);

                    queryRowsRequest.RopId = 0x15;
                    queryRowsRequest.LogonId = 0x00;
                    queryRowsRequest.InputHandleIndex = 0x00;
                    queryRowsRequest.QueryRowsFlags = 0x00;
                    queryRowsRequest.ForwardRead = 0x01;
                    queryRowsRequest.RowCount = 0x0032;
                    this.rowCountInQueryRowsRequest = 0x0032;
                    this.responseSOHs = this.DoSingleCallROP(queryRowsRequest, this.tableHandle, ref this.response, ref this.rawData);
                    queryRowsResponse = (RopQueryRowsResponse)this.response;

                    if (queryRowsResponse.RowData.PropertyRows == null)
                    {
                        Site.Assert.Fail("Current user should have received email messages before running this test case.");
                    }

                    ulong messageId = BitConverter.ToUInt64(queryRowsResponse.RowData.PropertyRows[0].PropertyValues[0].Value, 0);

                    RopOpenMessageRequest openMessageRequest;
                    RopOpenMessageResponse openMessageResponse;
                    openMessageRequest.RopId = 0x03;
                    openMessageRequest.LogonId = 0x0;
                    openMessageRequest.InputHandleIndex = 0x0;
                    openMessageRequest.OutputHandleIndex = 0x01;
                    openMessageRequest.CodePageId = 0x0FFF;
                    openMessageRequest.FolderId = this.folderID;
                    openMessageRequest.OpenModeFlags = 0x03; // Standard: Open for read/write if the user has write permissions for the folder, read-only if not.
                    openMessageRequest.MessageId = messageId;

                    this.responseSOHs = this.DoSingleCallROP(openMessageRequest, this.inputObjHandle, ref this.response, ref this.rawData);
                    openMessageResponse = (RopOpenMessageResponse)this.response;
                    uint messageObjHandle = this.responseSOHs[0][openMessageResponse.OutputHandleIndex];

                    RopGetAttachmentTableRequest getAttachmentTableRequest;
                    RopGetAttachmentTableResponse getAttachmentTableResponse;
                    getAttachmentTableRequest.RopId = 0x21;
                    getAttachmentTableRequest.LogonId = 0x00;
                    getAttachmentTableRequest.InputHandleIndex = 0x00;
                    getAttachmentTableRequest.OutputHandleIndex = 0x01;
                    getAttachmentTableRequest.TableFlags = 0x00; // Standard: Open the table.
                    this.responseSOHs = this.DoSingleCallROP(getAttachmentTableRequest, messageObjHandle, ref this.response, ref this.rawData);
                    getAttachmentTableResponse = (RopGetAttachmentTableResponse)this.response;
                    this.tableHandle = this.responseSOHs[0][getAttachmentTableResponse.OutputHandleIndex];
                    break;

                case TableType.HIERARCHY_TABLE:
                    #region Create [TestGetHierarchyTableFolder1] folder under Inbox
                    createFolderRequest.RopId = 0x1C;
                    createFolderRequest.LogonId = 0x00;
                    createFolderRequest.InputHandleIndex = 0x00;
                    createFolderRequest.OutputHandleIndex = 0x01;
                    createFolderRequest.FolderType = 0x01; // Generic folder
                    createFolderRequest.UseUnicodeStrings = 0x0; // FALSE
                    createFolderRequest.OpenExisting = 0x0; // FALSE
                    createFolderRequest.Reserved = 0x0; // FALSE
                    createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.TestGetHierarchyTableFolderName1 + "\0");
                    createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.TestGetHierarchyTableFolderName1 + "\0");

                    this.responseSOHs = this.DoSingleCallROP(createFolderRequest, this.folderHandle, ref this.response, ref this.rawData);
                    createFolderResponse = (RopCreateFolderResponse)this.response;
                    uint testGetHierarchyTableFolder1 = this.responseSOHs[0][createFolderResponse.OutputHandleIndex];
                    #endregion

                    #region Create [TestGetHierarchyTableFolder2] folder under [TestGetHierarchyTableFolder1]
                    createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.TestGetHierarchyTableFolderName2 + "\0");
                    createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.TestGetHierarchyTableFolderName2 + "\0");
                    this.responseSOHs = this.DoSingleCallROP(createFolderRequest, testGetHierarchyTableFolder1, ref this.response, ref this.rawData);
                    createFolderResponse = (RopCreateFolderResponse)this.response;
                    #endregion

                    RopGetHierarchyTableRequest getHierarchyTableRequest;
                    RopGetHierarchyTableResponse getHierarchyTableResponse;

                    getHierarchyTableRequest.RopId = 0x04;
                    getHierarchyTableRequest.LogonId = 0x00;
                    getHierarchyTableRequest.InputHandleIndex = 0x00;
                    getHierarchyTableRequest.OutputHandleIndex = 0x01;
                    getHierarchyTableRequest.TableFlags = 0x04; // Standard: The hierarchy table lists folders from all levels under the folder.

                    this.responseSOHs = this.DoSingleCallROP(getHierarchyTableRequest, this.folderHandle, ref this.response, ref this.rawData);
                    getHierarchyTableResponse = (RopGetHierarchyTableResponse)this.response;
                    this.tableHandle = this.responseSOHs[0][getHierarchyTableResponse.OutputHandleIndex];
                    this.rowCount = getHierarchyTableResponse.RowCount;
                    Site.Assert.AreEqual<uint>(0x02, this.rowCount, "The folder count in the inbox folder should be correct.");

                    break;
                case TableType.PERMISSIONS_TABLE:
                    RopGetPermissionsTableRequest getPermissionsTableRequest;
                    RopGetPermissionsTableResponse getPermissionsTableResponse;

                    getPermissionsTableRequest.RopId = 0x3E;
                    getPermissionsTableRequest.LogonId = 0x00;
                    getPermissionsTableRequest.InputHandleIndex = 0x00;
                    getPermissionsTableRequest.OutputHandleIndex = 0x01;
                    getPermissionsTableRequest.TableFlags = 0x00;

                    this.responseSOHs = this.DoSingleCallROP(getPermissionsTableRequest, this.folderHandle, ref this.response, ref this.rawData);
                    getPermissionsTableResponse = (RopGetPermissionsTableResponse)this.response;
                    this.tableHandle = this.responseSOHs[0][getPermissionsTableResponse.OutputHandleIndex];

                    break;
                case TableType.RULES_TABLE:
                    #region Create [testRuleTableFolder] folder under Inbox
                    createFolderRequest.RopId = 0x1C;
                    createFolderRequest.LogonId = 0x00;
                    createFolderRequest.InputHandleIndex = 0x00;
                    createFolderRequest.OutputHandleIndex = 0x01;
                    createFolderRequest.FolderType = 0x01; // Generic folder
                    createFolderRequest.UseUnicodeStrings = 0x0; // FALSE
                    createFolderRequest.OpenExisting = 0x0; // non-zero(TRUE)
                    createFolderRequest.Reserved = 0x0; // FALSE
                    createFolderRequest.DisplayName = Encoding.ASCII.GetBytes("TestRuleTableFolder" + "\0");
                    createFolderRequest.Comment = Encoding.ASCII.GetBytes("TestRuleTableFolder" + "\0");

                    this.responseSOHs = this.DoSingleCallROP(createFolderRequest, this.folderHandle, ref this.response, ref this.rawData);
                    createFolderResponse = (RopCreateFolderResponse)this.response;
                    this.folderHandle = this.responseSOHs[0][createFolderResponse.OutputHandleIndex];
                    this.folderID = createFolderResponse.FolderId;
                    #endregion

                    #region Add two rules to the created [testRuleTableFolder]
                    RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(Constants.RuleNameMarkAsRead);
                    RuleData ruleForMarkRead = AdapterHelper.GenerateValidRuleData(ActionTypes.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, null, ruleProperties);
                    RuleData ruleWithXRuleStateFlag = AdapterHelper.GenerateValidRuleData(ActionTypes.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.X, null, ruleProperties);
                    RuleData[] ruleDatas = { ruleForMarkRead, ruleWithXRuleStateFlag };

                    RopModifyRulesRequest modifyRulesRequest;

                    modifyRulesRequest.RopId = 0x41;
                    modifyRulesRequest.LogonId = 0x0;
                    modifyRulesRequest.InputHandleIndex = 0x00;
                    modifyRulesRequest.ModifyRulesFlags = (byte)ModifyRuleFlag.Modify_ReplaceAll;
                    modifyRulesRequest.RulesCount = 0;
                    if (ruleDatas != null)
                    {
                        modifyRulesRequest.RulesCount = (ushort)ruleDatas.Length;
                    }

                    modifyRulesRequest.RulesData = ruleDatas;

                    this.responseSOHs = this.DoSingleCallROP(modifyRulesRequest, this.folderHandle, ref this.response, ref this.rawData);
                    #endregion

                    RopGetRulesTableRequest getRulesTableRequest;
                    RopGetRulesTableResponse getRulesTableResponse;

                    getRulesTableRequest.RopId = 0x3F;
                    getRulesTableRequest.LogonId = 0x00;
                    getRulesTableRequest.InputHandleIndex = 0x00;
                    getRulesTableRequest.OutputHandleIndex = 0x01;
                    getRulesTableRequest.TableFlags = 0x00; // String values in the table are NOT in Unicode.

                    this.responseSOHs = this.DoSingleCallROP(getRulesTableRequest, this.folderHandle, ref this.response, ref this.rawData);
                    getRulesTableResponse = (RopGetRulesTableResponse)this.response;
                    this.tableHandle = this.responseSOHs[0][getRulesTableResponse.OutputHandleIndex];
                    tempPropertyTags = this.CreateSampleContentsTablePropertyTags();
                    setColumnsRequest.RopId = 0x12;
                    setColumnsRequest.LogonId = 0x00;
                    setColumnsRequest.InputHandleIndex = 0x00;
                    setColumnsRequest.SetColumnsFlags = 0x00;
                    setColumnsRequest.PropertyTagCount = (ushort)tempPropertyTags.Length;
                    setColumnsRequest.PropertyTags = tempPropertyTags;

                    this.responseSOHs = this.DoSingleCallROP(setColumnsRequest, this.tableHandle, ref this.response, ref this.rawData);
                    queryRowsRequest.RopId = 0x15;
                    queryRowsRequest.LogonId = 0x00;
                    queryRowsRequest.InputHandleIndex = 0x00;
                    queryRowsRequest.QueryRowsFlags = 0x01;
                    queryRowsRequest.ForwardRead = 0x01;
                    queryRowsRequest.RowCount = 0x0032;
                    this.rowCountInQueryRowsRequest = 0x0032;
                    this.responseSOHs = this.DoSingleCallROP(queryRowsRequest, this.tableHandle, ref this.response, ref this.rawData);
                    queryRowsResponse = (RopQueryRowsResponse)this.response;
                    this.rowCount = queryRowsResponse.RowCount;
                    Site.Assert.AreEqual<uint>(0x02, this.rowCount, "The rule count should be correct.");
                    break;
                case TableType.INVALID_TABLE:
                    this.tableHandle = this.folderHandle;

                    break;
            }
        }

        /// <summary>
        /// This methods is used to prepare the table with certain table type
        /// </summary>
        /// <param name="tableType">Identify the table type</param>
        private void PrepareTable(TableType tableType)
        {
            // Call open folder to prepare the table
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x0;
            openFolderRequest.InputHandleIndex = 0x0;
            openFolderRequest.OutputHandleIndex = 0x01;
            openFolderRequest.FolderId = this.folderID;
            openFolderRequest.OpenModeFlags = 0x0; // Opening an existing folder.

            this.responseSOHs = this.DoSingleCallROP(openFolderRequest, this.inputObjHandle, ref this.response, ref this.rawData);
            openFolderResponse = (RopOpenFolderResponse)this.response;

            this.folderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];
            this.inboxFolderHandle = this.folderHandle;

            // Get table handle
            this.GetTableHandle(tableType);
        }

        /// <summary>
        /// The method is used to verify whether the mail count is equal to the mails sent. 
        /// </summary>
        /// <param name="mailCount">The count of mails have been sent</param>
        private void VerifyMailCount(int mailCount)
        {
            // Verify row count and waiting it affect
            RopGetContentsTableRequest getContentsTableRequest;
            RopGetContentsTableResponse getContentsTableResponse;

            getContentsTableRequest.RopId = 0x05;
            getContentsTableRequest.LogonId = 0x00;
            getContentsTableRequest.InputHandleIndex = 0x00;
            getContentsTableRequest.OutputHandleIndex = 0x01;
            getContentsTableRequest.TableFlags = 0x00; // Depth.

            int getMailRetryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            while (getMailRetryCount >= 0)
            {
                this.responseSOHs = this.DoSingleCallROP(getContentsTableRequest, this.inboxFolderHandle, ref this.response, ref this.rawData);
                getContentsTableResponse = (RopGetContentsTableResponse)this.response;
                if (getContentsTableResponse.RowCount == mailCount)
                {
                    break;
                }

                getMailRetryCount--;
                Thread.Sleep(this.waitTime);
            }

            if (getMailRetryCount < 0)
            {
                Site.Assert.Fail("There should be {0} emails in the current user's inbox", mailCount);
            }
        }

        /// <summary>
        /// The method is used to clean up the inbox folder. 
        /// </summary>
        private void CleanInbox()
        {
            RopLogonResponse logonResponse = this.Logon(out this.inputObjHandle);
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = 0x02,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                OutputHandleIndex = 0x01,
                FolderId = logonResponse.FolderIds[4],
                OpenModeFlags = 0x00
            };
            this.responseSOHs = this.DoSingleCallROP(openFolderRequest, this.inputObjHandle, ref this.response, ref this.rawData);
            Site.Assert.AreEqual<uint>(0, ((RopOpenFolderResponse)this.response).ReturnValue, "Return value should be 0");
            this.inboxFolderHandle = this.responseSOHs[0][((RopOpenFolderResponse)this.response).OutputHandleIndex];

            RopHardDeleteMessagesAndSubfoldersRequest hardDeleteMessagesAndSubfoldersRequest;
            hardDeleteMessagesAndSubfoldersRequest.RopId = 0x92;
            hardDeleteMessagesAndSubfoldersRequest.LogonId = 0x00;
            hardDeleteMessagesAndSubfoldersRequest.InputHandleIndex = 0x00;
            hardDeleteMessagesAndSubfoldersRequest.WantAsynchronous = 0x00; // Synchronously
            hardDeleteMessagesAndSubfoldersRequest.WantDeleteAssociated = 0xFF; // TRUE: delete all messages and subfolders
            object temp = new object();
            this.DoSingleCallROP(hardDeleteMessagesAndSubfoldersRequest, this.inboxFolderHandle, ref temp, ref this.rawData);
            Site.Assert.AreEqual<uint>(0, ((RopHardDeleteMessagesAndSubfoldersResponse)temp).ReturnValue, "Return value should be 0");

            // Waiting for the server to solve the delete request.
            Thread.Sleep(this.waitTime);
            this.VerifyMailCount(0);
        }

        /// <summary>
        /// This method is used to check whether the sender's name is correct.
        /// </summary>
        /// <param name="actualSenderName">The value of the PidTagSenderName property returned by server</param>
        /// <param name="expectedSenderName">The user name preconfigured in PTFConfig file</param>
        /// <returns>Identify whether it is expected Sender</returns>
        private bool IsExpectedSender(string actualSenderName, string expectedSenderName)
        {
            bool isExpectedSender = false;

            // Under Exchange 2013, the sender name contains the domain portion, which begins with '@'.
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            int domainIndex = actualSenderName.IndexOf("@" + domain, StringComparison.CurrentCultureIgnoreCase);
            if (domainIndex > 0)
            {
                actualSenderName = actualSenderName.Remove(domainIndex, domain.Length + 1);
            }

            // As specified in [MS-OXOMSG] section 2.2.1.48 and [MS-OXCDATA] section 2.8.3.2, the sender name should be a null-terminated string.
            expectedSenderName = expectedSenderName + "\0";
            isExpectedSender = string.Equals(actualSenderName, expectedSenderName, StringComparison.CurrentCultureIgnoreCase);
            return isExpectedSender;
        }

        /// <summary>
        /// This method is used to create message as the content of the content table or attachment table.
        /// </summary>
        /// <param name="senderName">The user name used as the sender of the message to be created</param>
        private void CreateMessageForPrepareTable(string senderName)
        {
            #region Create a message
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;
            createMessageRequest.RopId = 0x06;
            createMessageRequest.LogonId = 0x0;
            createMessageRequest.InputHandleIndex = 0x0;
            createMessageRequest.OutputHandleIndex = 0x01;
            createMessageRequest.CodePageId = 0x0FFF; // Code page of Logon object is used
            createMessageRequest.FolderId = this.folderID;
            createMessageRequest.AssociatedFlag = 0x00; // NOT an FAI message
            this.responseSOHs = this.DoSingleCallROP(createMessageRequest, this.inputObjHandle, ref this.response, ref this.rawData);
            createMessageResponse = (RopCreateMessageResponse)this.response;
            uint messageHandle = this.responseSOHs[0][createMessageResponse.OutputHandleIndex];
            Site.Assert.AreEqual<uint>(
                0,
                createMessageResponse.ReturnValue,
                "0 indicates the RopCreateMessage ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Set message properties
            RopSetPropertiesRequest setPropertiesRequest;
            RopSetPropertiesResponse setPropertiesResponse;

            setPropertiesRequest.RopId = 0x0A;
            setPropertiesRequest.LogonId = 0x0;
            setPropertiesRequest.InputHandleIndex = 0x00;
            TaggedPropertyValue[] taggedPropertyValueArray = new TaggedPropertyValue[2];

            // PidTagSenderName
            taggedPropertyValueArray[0] = new TaggedPropertyValue();
            PropertyTag pidTagMSenderName = new PropertyTag
            {
                PropertyId = 0x0C1A,
                PropertyType = (ushort)PropertyTypeName.PtypString
            };
            taggedPropertyValueArray[0].PropertyTag = pidTagMSenderName;
            taggedPropertyValueArray[0].Value = Encoding.Unicode.GetBytes(senderName + "\0");

            // PidTagMessageDeliveryTime
            taggedPropertyValueArray[1] = new TaggedPropertyValue();
            PropertyTag pidTagMessageDeliveryTime = new PropertyTag
            {
                PropertyId = 0x0e06, PropertyType = 0x0040
            };
            taggedPropertyValueArray[1].PropertyTag = pidTagMessageDeliveryTime;
            long dateTime = DateTime.Now.ToFileTime();
            taggedPropertyValueArray[1].Value = BitConverter.GetBytes(dateTime);
            if (taggedPropertyValueArray != null)
            {
                setPropertiesRequest.PropertyValueCount = (ushort)taggedPropertyValueArray.Length;
                ushort count = 2;
                foreach (TaggedPropertyValue tagValue in taggedPropertyValueArray)
                {
                    count += (ushort)(tagValue.Value.Length + 4);
                }

                setPropertiesRequest.PropertyValueSize = count;
            }
            else
            {
                setPropertiesRequest.PropertyValueCount = 0x00;
                setPropertiesRequest.PropertyValueSize = 2;
            }

            setPropertiesRequest.PropertyValues = taggedPropertyValueArray;
            this.responseSOHs = this.DoSingleCallROP(setPropertiesRequest, messageHandle, ref this.response, ref this.rawData);
            setPropertiesResponse = (RopSetPropertiesResponse)this.response;
            Site.Assert.AreEqual<uint>(
                0,
                setPropertiesResponse.ReturnValue,
                "0 indicates the RopSetProperties ROP succeeds, other value indicates error occurs.");

            #endregion

            #region Save the message
            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageRequest.RopId = 0x0C;
            saveChangesMessageRequest.LogonId = 0x0;
            saveChangesMessageRequest.InputHandleIndex = 0x0;
            saveChangesMessageRequest.ResponseHandleIndex = 0x01;
            saveChangesMessageRequest.SaveFlags = 0x0C; // ForceSave
            this.responseSOHs = this.DoSingleCallROP(saveChangesMessageRequest, messageHandle, ref this.response, ref this.rawData);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(
                0,
                saveChangesMessageResponse.ReturnValue,
                "0 indicates the RopSaveChangesMessage ROP succeeds, other value indicates error occurs.");
            #endregion
        }
        #endregion
    }
}