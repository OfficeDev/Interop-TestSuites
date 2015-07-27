//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The MapiHttpAdapter class contains the MAPIHTTP implements for the interfaces of IMS_OXNSPIAdapter.
    /// </summary>
    public class NspiMapiHttpAdapter
    {
        #region Variables

        /// <summary>
        /// The Site instance.
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// The Mailbox userName which can be used by client to connect to the SUT.
        /// </summary>
        private string userName;

        /// <summary>
        /// The user password which can be used by client to access to the SUT.
        /// </summary>
        private string password;

        /// <summary>
        /// Define the name of domain where the server belongs to.
        /// </summary>
        private string domainName;

        /// <summary>
        /// The URL that a client can use to connect with a NSPI server through MAPI over HTTP.
        /// </summary>
        private string addressBookUrl;

        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="NspiMapiHttpAdapter" /> class.
        /// </summary>
        /// <param name="site">The Site instance.</param>
        /// <param name="userName">The Mailbox userName which can be used by client to connect to the SUT.</param>
        /// <param name="password">The user password which can be used by client to access to the SUT.</param>
        /// <param name="domainName">Define the name of domain where the server belongs to.</param>
        /// <param name="addressBookUrl">The URL that a client can use to connect with a NSPI server through MAPI over HTTP.</param>
        public NspiMapiHttpAdapter(ITestSite site, string userName, string password, string domainName, string addressBookUrl)
        {
            this.site = site;
            this.userName = userName;
            this.password = password;
            this.domainName = domainName;
            this.addressBookUrl = addressBookUrl;
        }

        #region Instance interface

        /// <summary>
        /// The NspiBind method initiates a session between a client and the server.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="serverGuid">The value NULL or a pointer to a GUID value that is associated with the specific server.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue Bind(uint flags, STAT stat, ref FlatUID_r? serverGuid)
        {
            ErrorCodeValue result;
            BindRequestBody bindRequestBody = this.BuildBindRequestBody(stat, flags);
            byte[] rawBuffer = null;
            ChunkedResponse chunkedResponse = null;
            BindResponseBody bindResponseBody = null;

            // Send the execute HTTP request and get the response
            HttpWebResponse response = MapiHttpAdapter.SendMAPIHttpRequest(this.site, this.addressBookUrl, this.userName, this.domainName, this.password, bindRequestBody, RequestType.Bind.ToString(), AdapterHelper.SessionContextCookies);

            // Read the HTTP response buffer and parse the response to correct format
            rawBuffer = MapiHttpAdapter.ReadHttpResponse(response);
            result = (ErrorCodeValue)int.Parse(response.Headers["X-ResponseCode"]);
            if (result == ErrorCodeValue.Success)
            {
                chunkedResponse = ChunkedResponse.ParseChunkedResponse(rawBuffer);
                bindResponseBody = BindResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
                result = (ErrorCodeValue)bindResponseBody.ErrorCode;
                if (bindResponseBody.ServerGuid != null)
                {
                    FlatUID_r newGuid = new FlatUID_r();
                    newGuid.Ab = bindResponseBody.ServerGuid.ToByteArray();
                    serverGuid = newGuid;
                }
                else
                {
                    serverGuid = null;
                }
            }

            response.GetResponseStream().Close();
            AdapterHelper.SessionContextCookies = response.Cookies;
            return result;
        }

        /// <summary>
        /// The NspiUnbind method destroys the context handle. No other action is taken.
        /// </summary>
        /// <param name="reserved">A DWORD [MS-DTYP] value reserved for future use. This property is ignored by the server.</param>
        /// <returns>A DWORD value that specifies the return status of the method.</returns>
        public uint Unbind(uint reserved)
        {
            uint result;
            UnbindRequestBody unbindRequest = this.BuildUnbindRequestBody();
            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(unbindRequest, RequestType.Unbind);
            AdapterHelper.SessionContextCookies = new CookieCollection();
            UnbindResponseBody unbindResponseBody = UnbindResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = unbindResponseBody.ErrorCode;

            return result;
        }

        /// <summary>
        /// The NspiGetSpecialTable method returns the rows of a special table to the client. 
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="version">A reference to a DWORD. On input, it holds the value of the version number of
        /// the address book hierarchy table that the client has. On output, it holds the version of the server's address book hierarchy table.</param>
        /// <param name="rows">A PropertyRowSet_r structure. On return, it holds the rows for the table that the client is requesting.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue GetSpecialTable(uint flags, ref STAT stat, ref uint version, out PropertyRowSet_r? rows)
        {
            ErrorCodeValue result;
            byte[] auxIn = new byte[] { };
            GetSpecialTableRequestBody getSpecialTableRequestBody = new GetSpecialTableRequestBody()
            {
                Flags = flags,
                HasState = true,
                State = stat,
                HasVersion = true,
                Version = version,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(getSpecialTableRequestBody, RequestType.GetSpecialTable);
            GetSpecialTableResponseBody getSpecialTableResponseBody = GetSpecialTableResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)getSpecialTableResponseBody.ErrorCode;
            if (getSpecialTableResponseBody.HasRows)
            {
                PropertyRowSet_r newRows = AdapterHelper.ParsePropertyRowSet_r(getSpecialTableResponseBody.RowCount.Value, getSpecialTableResponseBody.Rows);
                rows = newRows;
            }
            else
            {
                rows = null;
            }

            if (getSpecialTableResponseBody.HasVersion)
            {
                version = getSpecialTableResponseBody.Version.Value;
            }

            return result;
        }

        /// <summary>
        /// The NspiUpdateStat method updates the STAT block that represents the position in a table 
        /// to reflect positioning changes requested by the client.
        /// </summary>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="delta">The value NULL or a pointer to a LONG value that indicates movement 
        /// within the address book container specified by the input parameter stat.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue UpdateStat(ref STAT stat, ref int? delta)
        {
            ErrorCodeValue result;
            UpdateStatRequestBody updateStatRequestBody = this.BuildUpdateStatRequestBody(stat);
            if (delta == null)
            {
                updateStatRequestBody.DeltaRequested = false;
            }

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(updateStatRequestBody, RequestType.UpdateStat);
            UpdateStatResponseBody updateStatResponseBody = UpdateStatResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)updateStatResponseBody.ErrorCode;
            if (updateStatResponseBody.HasDelta)
            {
                delta = updateStatResponseBody.Delta;
            }

            if (updateStatResponseBody.HasState)
            {
                stat = updateStatResponseBody.State.Value;
            }

            return result;
        }

        /// <summary>
        /// The NspiQueryColumns method returns a list of all the properties that the server is aware of. 
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="columns">A PropertyTagArray_r structure that contains a list of proptags.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue QueryColumns(uint flags, out PropertyTagArray_r? columns)
        {
            ErrorCodeValue result;
            QueryColumnsRequestBody queryColumnsRequestBody = this.BuildQueryColumnsRequestBody(flags);
            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(queryColumnsRequestBody, RequestType.QueryColumns);
            QueryColumnsResponseBody queryColumnsResponseBody = QueryColumnsResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)queryColumnsResponseBody.ErrorCode;
            if (queryColumnsResponseBody.HasColumns)
            {
                PropertyTagArray_r propertyTagArray = new PropertyTagArray_r();
                propertyTagArray.CValues = queryColumnsResponseBody.Columns.Value.PropertyTagCount;
                propertyTagArray.AulPropTag = new uint[propertyTagArray.CValues];
                for (int i = 0; i < propertyTagArray.CValues; i++)
                {
                    propertyTagArray.AulPropTag[i] = (uint)((queryColumnsResponseBody.Columns.Value.PropertyTags[i].PropertyId * 65536) | queryColumnsResponseBody.Columns.Value.PropertyTags[i].PropertyType);
                }

                columns = propertyTagArray;
            }
            else
            {
                columns = null;
            }

            return result;
        }

        /// <summary>
        /// The NspiGetPropList method returns a list of all the properties that have values on a specified object.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="mid">A DWORD value that contains a Minimal Entry ID.</param>
        /// <param name="codePage">The code page in which the client wants the server to express string values properties.</param>
        /// <param name="propTags">A PropertyTagArray_r value. On return, it holds a list of properties.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue GetPropList(uint flags, uint mid, uint codePage, out PropertyTagArray_r? propTags)
        {
            ErrorCodeValue result;
            GetPropListRequestBody getPropListRequestBody = this.BuildGetPropListRequestBody(flags, mid, codePage);
            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(getPropListRequestBody, RequestType.GetPropList);
            GetPropListResponseBody getPropListResponseBody = GetPropListResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)getPropListResponseBody.ErrorCode;
            if (getPropListResponseBody.HasPropertyTags)
            {
                PropertyTagArray_r propertyTagArray = new PropertyTagArray_r();
                propertyTagArray.CValues = getPropListResponseBody.PropertyTags.Value.PropertyTagCount;
                propertyTagArray.AulPropTag = new uint[propertyTagArray.CValues];
                for (int i = 0; i < propertyTagArray.CValues; i++)
                {
                    propertyTagArray.AulPropTag[i] = (uint)((getPropListResponseBody.PropertyTags.Value.PropertyTags[i].PropertyId * 65536) | getPropListResponseBody.PropertyTags.Value.PropertyTags[i].PropertyType);
                }

                propTags = propertyTagArray;
            }
            else
            {
                propTags = null;
            }

            return result;
        }

        /// <summary>
        /// The NspiGetProps method returns an address book row that contains a set of the properties
        /// and values that exist on an object.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value. 
        /// It contains a list of the proptags of the properties that the client wants to be returned.</param>
        /// <param name="rows">A nullable PropertyRow_r value. 
        /// It contains the address book container row the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue GetProps(uint flags, STAT stat, PropertyTagArray_r? propTags, out PropertyRow_r? rows)
        {
            ErrorCodeValue result;
            GetPropsRequestBody getPropertyRequestBody = null;
            LargePropTagArray propetyTags = new LargePropTagArray();
            if (propTags != null)
            {
                propetyTags.PropertyTagCount = propTags.Value.CValues;
                propetyTags.PropertyTags = new PropertyTag[propetyTags.PropertyTagCount];
                for (int i = 0; i < propTags.Value.CValues; i++)
                {
                    propetyTags.PropertyTags[i].PropertyId = (ushort)((propTags.Value.AulPropTag[i] & 0xFFFF0000) >> 16);
                    propetyTags.PropertyTags[i].PropertyType = (ushort)(propTags.Value.AulPropTag[i] & 0x0000FFFF);
                }

                getPropertyRequestBody = this.BuildGetPropsRequestBody(flags, true, stat, true, propetyTags);
            }
            else
            {
                getPropertyRequestBody = this.BuildGetPropsRequestBody(flags, true, stat, false, propetyTags);
            }

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(getPropertyRequestBody, RequestType.GetProps);
            GetPropsResponseBody getPropsResponseBody = GetPropsResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)getPropsResponseBody.ErrorCode;
            if (getPropsResponseBody.HasPropertyValues)
            {
                PropertyRow_r propertyRow = AdapterHelper.ParsePropertyRow_r(getPropsResponseBody.PropertyValues.Value);
                rows = propertyRow;
            }
            else
            {
                rows = null;
            }

            return result;
        }

        /// <summary>
        /// The NspiQueryRows method returns a number of rows from a specified table to the client.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="tableCount">A DWORD value that contains the number values in the input parameter table. 
        /// This value is limited to 100,000.</param>
        /// <param name="table">An array of DWORD values, representing an Explicit Table.</param>
        /// <param name="count">A DWORD value that contains the number of rows the client is requesting.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value, 
        /// containing a list of the proptags of the properties that the client requires to be returned for each row returned.</param>
        /// <param name="rows">A nullable PropertyRowSet_r value, it contains the address book container rows that the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue QueryRows(uint flags, ref STAT stat, uint tableCount, uint[] table, uint count, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows)
        {
            ErrorCodeValue result;
            QueryRowsRequestBody queryRowsRequestBody = new QueryRowsRequestBody();
            LargePropTagArray propetyTags = new LargePropTagArray();
            if (propTags != null)
            {
                propetyTags.PropertyTagCount = propTags.Value.CValues;
                propetyTags.PropertyTags = new PropertyTag[propetyTags.PropertyTagCount];
                for (int i = 0; i < propTags.Value.CValues; i++)
                {
                    propetyTags.PropertyTags[i].PropertyId = (ushort)((propTags.Value.AulPropTag[i] & 0xFFFF0000) >> 16);
                    propetyTags.PropertyTags[i].PropertyType = (ushort)(propTags.Value.AulPropTag[i] & 0x0000FFFF);
                }

                queryRowsRequestBody.HasColumns = true;
                queryRowsRequestBody.Columns = propetyTags;
            }

            queryRowsRequestBody.Flags = flags;
            queryRowsRequestBody.HasState = true;
            queryRowsRequestBody.State = stat;
            queryRowsRequestBody.ExplicitTableCount = tableCount;
            queryRowsRequestBody.ExplicitTable = table;
            queryRowsRequestBody.RowCount = count;
            byte[] auxIn = new byte[] { };
            queryRowsRequestBody.AuxiliaryBuffer = auxIn;
            queryRowsRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(queryRowsRequestBody, RequestType.QueryRows);
            QueryRowsResponseBody queryRowsResponseBody = QueryRowsResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)queryRowsResponseBody.ErrorCode;
            if (queryRowsResponseBody.RowCount != null)
            {
                PropertyRowSet_r newRows = AdapterHelper.ParsePropertyRowSet_r(queryRowsResponseBody.Columns.Value, queryRowsResponseBody.RowCount.Value, queryRowsResponseBody.RowData);
                rows = newRows;
            }
            else
            {
                rows = null;
            }

            if (queryRowsResponseBody.HasState)
            {
                stat = queryRowsResponseBody.State.Value;
            }

            return result;
        }

        /// <summary>
        /// The NspiSeekEntries method searches for and sets the logical position in a specific table
        /// to the first entry greater than or equal to a specified value. 
        /// </summary>
        /// <param name="reserved">A DWORD value that is reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="target">A PropertyValue_r value holding the value that is being sought.</param>
        /// <param name="table">The value NULL or a PropertyTagArray_r value. 
        /// It holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="propTags">It contains a list of the proptags of the columns 
        /// that client wants to be returned for each row returned.</param>
        /// <param name="rows">It contains the address book container rows the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue SeekEntries(uint reserved, ref STAT stat, PropertyValue_r target, PropertyTagArray_r? table, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows)
        {
            ErrorCodeValue result;
            SeekEntriesRequestBody seekEntriesRequestBody = new SeekEntriesRequestBody();
            TaggedPropertyValue targetValue = new TaggedPropertyValue();
            targetValue.PropertyTag = new PropertyTag((ushort)((target.PropTag & 0xFFFF0000) >> 16), (ushort)(target.PropTag & 0x0000FFFF));
            targetValue.Value = new byte[target.Serialize().Length - 8];
            Array.Copy(target.Serialize(), 8, targetValue.Value, 0, target.Serialize().Length - 8);

            // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
            seekEntriesRequestBody.Reserved = reserved;
            seekEntriesRequestBody.HasState = true;
            seekEntriesRequestBody.State = stat;
            seekEntriesRequestBody.HasTarget = true;
            seekEntriesRequestBody.Target = targetValue;
            if (table != null)
            {
                seekEntriesRequestBody.HasExplicitTable = true;
                seekEntriesRequestBody.ExplicitTable = table.Value.AulPropTag;
            }

            LargePropTagArray propetyTags = new LargePropTagArray();
            if (propTags != null)
            {
                propetyTags.PropertyTagCount = propTags.Value.CValues;
                propetyTags.PropertyTags = new PropertyTag[propetyTags.PropertyTagCount];
                for (int i = 0; i < propTags.Value.CValues; i++)
                {
                    propetyTags.PropertyTags[i].PropertyId = (ushort)((propTags.Value.AulPropTag[i] & 0xFFFF0000) >> 16);
                    propetyTags.PropertyTags[i].PropertyType = (ushort)(propTags.Value.AulPropTag[i] & 0x0000FFFF);
                }

                seekEntriesRequestBody.HasColumns = true;
                seekEntriesRequestBody.Columns = propetyTags;
            }

            seekEntriesRequestBody.AuxiliaryBufferSize = 0;
            seekEntriesRequestBody.AuxiliaryBuffer = new byte[] { };
            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(seekEntriesRequestBody, RequestType.SeekEntries);
            SeekEntriesResponseBody seekEntriesResponseBody = SeekEntriesResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)seekEntriesResponseBody.ErrorCode;
            if (seekEntriesResponseBody.HasColumnsAndRows)
            {
                PropertyRowSet_r newRows = AdapterHelper.ParsePropertyRowSet_r(seekEntriesResponseBody.Columns.Value, seekEntriesResponseBody.RowCount.Value, seekEntriesResponseBody.RowData);
                rows = newRows;
            }
            else
            {
                rows = null;
            }

            if (seekEntriesResponseBody.HasState)
            {
                stat = seekEntriesResponseBody.State.Value;
            }

            return result;
        }

        /// <summary>
        /// The NspiGetMatches method returns an Explicit Table. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="proReserved">A PropertyTagArray_r reserved for future use.</param>
        /// <param name="reserved2">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="filter">The value NULL or a Restriction_r value. 
        /// It holds a logical restriction to apply to the rows in the address book container specified in the stat parameter.</param>
        /// <param name="propName">The value NULL or a PropertyName_r value. 
        /// It holds the property to be opened as a restricted address book container.</param>
        /// <param name="requested">A DWORD value. It contains the maximum number of rows to return in a restricted address book container.</param>
        /// <param name="outMids">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value. 
        /// It contains a list of the proptags of the columns that client wants to be returned for each row returned.</param>
        /// <param name="rows">A reference to a PropertyRowSet_r value. It contains the address book container rows the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue GetMatches(uint reserved, ref STAT stat, PropertyTagArray_r? proReserved, uint reserved2, Restriction_r? filter, PropertyName_r? propName, uint requested, out PropertyTagArray_r? outMids, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows)
        {
            ErrorCodeValue result;
            byte[] auxIn = new byte[] { };
            GetMatchesRequestBody getMatchesRequestBody = new GetMatchesRequestBody()
            {
                Reserved = reserved,
                HasState = true,
                State = stat,
                InterfaceOptionFlags = reserved2,
                HasPropertyName = false,
                RowCount = requested,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            if (propTags != null)
            {
                LargePropTagArray propetyTags = new LargePropTagArray();
                propetyTags.PropertyTagCount = propTags.Value.CValues;
                propetyTags.PropertyTags = new PropertyTag[propetyTags.PropertyTagCount];
                for (int i = 0; i < propTags.Value.CValues; i++)
                {
                    propetyTags.PropertyTags[i].PropertyId = (ushort)((propTags.Value.AulPropTag[i] & 0xFFFF0000) >> 16);
                    propetyTags.PropertyTags[i].PropertyType = (ushort)(propTags.Value.AulPropTag[i] & 0x0000FFFF);
                }

                getMatchesRequestBody.HasColumns = true;
                getMatchesRequestBody.Columns = propetyTags;
            }

            if (proReserved != null)
            {
                getMatchesRequestBody.HasMinimalIds = true;
                getMatchesRequestBody.MinimalIdCount = proReserved.Value.CValues;
                getMatchesRequestBody.MinimalIds = proReserved.Value.AulPropTag;
            }

            if (filter != null)
            {
                getMatchesRequestBody.HasFilter = true;
                getMatchesRequestBody.Filter = AdapterHelper.ConvertRestriction_rToRestriction(filter.Value);
            }

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(getMatchesRequestBody, RequestType.GetMatches);
            GetMatchesResponseBody getMatchesResponseBody = GetMatchesResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)getMatchesResponseBody.ErrorCode;
            if (getMatchesResponseBody.HasMinimalIds)
            {
                PropertyTagArray_r propertyTagArray = new PropertyTagArray_r();
                propertyTagArray.CValues = getMatchesResponseBody.MinimalIdCount.Value;
                propertyTagArray.AulPropTag = getMatchesResponseBody.MinimalIds;
                outMids = propertyTagArray;
            }
            else
            {
                outMids = null;
            }

            if (getMatchesResponseBody.RowCount != null)
            {
                PropertyRowSet_r newRows = AdapterHelper.ParsePropertyRowSet_r(getMatchesResponseBody.Columns.Value, getMatchesResponseBody.RowCount.Value, getMatchesResponseBody.RowData);
                rows = newRows;
            }
            else
            {
                rows = null;
            }

            if (getMatchesResponseBody.HasState)
            {
                stat = getMatchesResponseBody.State.Value;
            }

            return result;
        }

        /// <summary>
        /// The NspiResortRestriction method applies to a sort order to the objects in a restricted address book container.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="proInMIds">A PropertyTagArray_r value. 
        /// It holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="outMIds">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs 
        /// that comprise a restricted address book container.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue ResortRestriction(uint reserved, ref STAT stat, PropertyTagArray_r proInMIds, ref PropertyTagArray_r? outMIds)
        {
            ErrorCodeValue result;
            byte[] auxIn = new byte[] { };
            ResortRestrictionRequestBody resortRestrictionRequestBody = new ResortRestrictionRequestBody()
            {
                // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
                Reserved = reserved,
                HasState = true,
                State = stat,
                HasMinimalIds = true,
                MinimalIdCount = proInMIds.CValues,
                MinimalIds = proInMIds.AulPropTag,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(resortRestrictionRequestBody, RequestType.ResortRestriction);
            ResortRestrictionResponseBody resortRestrictionResponseBody = ResortRestrictionResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)resortRestrictionResponseBody.ErrorCode;
            if (resortRestrictionResponseBody.HasMinimalIds)
            {
                PropertyTagArray_r propertyTagArray = AdapterHelper.ParsePropertyTagArray_r(resortRestrictionResponseBody.MinimalIdCount.Value, resortRestrictionResponseBody.MinimalIds);
                outMIds = propertyTagArray;
            }
            else
            {
                outMIds = null;
            }

            if (resortRestrictionResponseBody.HasState)
            {
                stat = resortRestrictionResponseBody.State.Value;
            }

            return result;
        }

        /// <summary>
        /// The NspiCompareMIds method compares the position in an address book container of two objects 
        /// identified by Minimal Entry ID and returns the value of the comparison.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="mid1">The mid1 is a DWORD value containing a Minimal Entry ID.</param>
        /// <param name="mid2">The mid2 is a DWORD value containing a Minimal Entry ID.</param>
        /// <param name="results">A DWORD value. On return, it contains the result of the comparison.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue CompareMIds(uint reserved, STAT stat, uint mid1, uint mid2, out int results)
        {
            ErrorCodeValue result;
            byte[] auxIn = new byte[] { };
            CompareMinIdsRequestBody compareMinIdsRequestBody = new CompareMinIdsRequestBody()
            {
                // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
                Reserved = reserved,
                HasState = true,
                State = stat,
                MinimalId1 = mid1,
                MinimalId2 = mid2,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(compareMinIdsRequestBody, RequestType.CompareMIds);
            CompareMinIdsResponseBody compareMinIdsResponseBody = CompareMinIdsResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)compareMinIdsResponseBody.ErrorCode;
            results = compareMinIdsResponseBody.Result;
            return result;
        }

        /// <summary>
        /// The NspiDNToMId method maps a set of DN to a set of Minimal Entry ID.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="names">A StringsArray_r value. It holds a list of strings that contain DNs.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue DNToMId(uint reserved, StringsArray_r names, out PropertyTagArray_r? mids)
        {
            ErrorCodeValue result;
            byte[] auxIn = new byte[] { };
            StringArray_r nameArray = new StringArray_r();
            nameArray.CValues = names.CValues;
            nameArray.LppszA = names.LppszA;
            DNToMinIdRequestBody requestBodyOfdnToMId = new DNToMinIdRequestBody()
            {
                Reserved = reserved,
                HasNames = true,
                Names = nameArray,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(requestBodyOfdnToMId, RequestType.DNToMId);
            DnToMinIdResponseBody distinguishedNameToMinIdResponseBody = DnToMinIdResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)distinguishedNameToMinIdResponseBody.ErrorCode;
            if (distinguishedNameToMinIdResponseBody.HasMinimalIds)
            {
                PropertyTagArray_r propertyTagArray = AdapterHelper.ParsePropertyTagArray_r(distinguishedNameToMinIdResponseBody.MinimalIdCount.Value, distinguishedNameToMinIdResponseBody.MinimalIds);
                mids = propertyTagArray;
            }
            else
            {
                mids = null;
            }

            return result;
        }

        /// <summary>
        /// The NspiModProps method is used to modify the properties of an object in the address book. 
        /// </summary>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r. 
        /// It contains a list of the proptags of the columns from which the client requests all the values to be removed.</param>
        /// <param name="row">A PropertyRow_r value. It contains an address book row.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue ModProps(STAT stat, PropertyTagArray_r? propTags, PropertyRow_r row)
        {
            ErrorCodeValue result;
            ModPropsRequestBody modPropsRequestBody = new ModPropsRequestBody();
            modPropsRequestBody.HasState = true;
            modPropsRequestBody.State = stat;
            if (propTags != null)
            {
                LargePropTagArray largePropTagArray = new LargePropTagArray();
                largePropTagArray.PropertyTagCount = propTags.Value.CValues;
                largePropTagArray.PropertyTags = new PropertyTag[propTags.Value.CValues];
                for (int i = 0; i < propTags.Value.CValues; i++)
                {
                    largePropTagArray.PropertyTags[i].PropertyId = (ushort)((propTags.Value.AulPropTag[i] & 0xFFFF0000) >> 16);
                    largePropTagArray.PropertyTags[i].PropertyType = (ushort)(propTags.Value.AulPropTag[i] & 0x0000FFFF);
                }

                modPropsRequestBody.HasPropertyTagsToRemove = true;
                modPropsRequestBody.PropertyTagsToRemove = largePropTagArray;
            }
            else
            {
                modPropsRequestBody.HasPropertyTagsToRemove = false;
            }

            modPropsRequestBody.HasPropertyValues = true;
            AddressBookPropValueList addressBookPropValueList = new AddressBookPropValueList();
            addressBookPropValueList.PropertyValueCount = row.CValues;
            addressBookPropValueList.PropertyValues = new TaggedPropertyValue[row.CValues];
            for (int i = 0; i < row.CValues; i++)
            {
                addressBookPropValueList.PropertyValues[i] = new TaggedPropertyValue();
                byte[] propertyBytes = new byte[row.LpProps[i].Serialize().Length - 8];
                Array.Copy(row.LpProps[i].Serialize(), 8, propertyBytes, 0, row.LpProps[i].Serialize().Length - 8);
                addressBookPropValueList.PropertyValues[i].Value = propertyBytes;
                PropertyTag propertyTagOfRow = new PropertyTag();
                propertyTagOfRow.PropertyId = (ushort)((row.LpProps[i].PropTag & 0xFFFF0000) >> 16);
                propertyTagOfRow.PropertyType = (ushort)(row.LpProps[i].PropTag & 0x0000FFFF);
                addressBookPropValueList.PropertyValues[i].PropertyTag = propertyTagOfRow;
            }

            modPropsRequestBody.PropertyVaules = addressBookPropValueList;

            byte[] auxIn = new byte[] { };
            modPropsRequestBody.AuxiliaryBuffer = auxIn;
            modPropsRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(modPropsRequestBody, RequestType.ModProps);
            ModPropsResponseBody modPropsResponseBody = ModPropsResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)modPropsResponseBody.ErrorCode;
            return result;
        }

        /// <summary>
        /// The NspiModLinkAtt method modifies the values of a specific property of a specific row in the address book.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="propTag">A DWORD value. It contains the proptag of the property that the client wants to modify.</param>
        /// <param name="mid">A DWORD value that contains the Minimal Entry ID of the address book row that the client wants to modify.</param>
        /// <param name="entryIds">A BinaryArray value. It contains a list of EntryIDs to be used to modify the requested property on the requested address book row.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue ModLinkAtt(uint flags, uint propTag, uint mid, BinaryArray_r entryIds)
        {
            ErrorCodeValue result;
            ModLinkAttRequestBody modLinkAttRequestBody = new ModLinkAttRequestBody();
            modLinkAttRequestBody.Flags = flags;
            modLinkAttRequestBody.PropertyTag = new PropertyTag()
            {
                PropertyId = (ushort)((propTag & 0xFFFF0000) >> 16),
                PropertyType = (ushort)(propTag & 0x0000FFFF)
            };
            modLinkAttRequestBody.MinimalId = mid;
            if (entryIds.CValues != 0)
            {
                modLinkAttRequestBody.HasEntryIds = true;
                modLinkAttRequestBody.EntryIdCount = entryIds.CValues;
                modLinkAttRequestBody.EntryIDs = new byte[entryIds.CValues][];
                for (int i = 0; i < entryIds.CValues; i++)
                {
                    List<byte> entryIDBytes = new List<byte>();
                    entryIDBytes.AddRange(BitConverter.GetBytes((uint)entryIds.Lpbin[i].Lpb.Length));
                    entryIDBytes.AddRange(entryIds.Lpbin[i].Lpb);
                    modLinkAttRequestBody.EntryIDs[i] = entryIDBytes.ToArray();
                }
            }
            else
            {
                modLinkAttRequestBody.HasEntryIds = false;
            }

            byte[] auxIn = new byte[] { };
            modLinkAttRequestBody.AuxiliaryBuffer = auxIn;
            modLinkAttRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(modLinkAttRequestBody, RequestType.ModLinkAtt);
            ModLinkAttResponseBody modLinkAttResponseBody = ModLinkAttResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)modLinkAttResponseBody.ErrorCode;
            return result;
        }

        /// <summary>
        /// The NspiResolveNamesW method takes a set of string values in the Unicode character set 
        /// and performs ANR on those strings. 
        /// </summary>
        /// <param name="reserved">A DWORD value that is reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r containing a list of the proptags of the columns 
        /// that the client requests to be returned for each row returned.</param>
        /// <param name="wstr">A WStringsArray_r value. It specifies the values on which the client is requesting the server to perform ANR.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it contains a list of Minimal Entry IDs that match the array of strings.</param>
        /// <param name="rows">A reference to a PropertyRowSet_r structure. It contains the address book container rows that the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue ResolveNames(uint reserved, STAT stat, PropertyTagArray_r? propTags, WStringsArray_r? wstr, out PropertyTagArray_r? mids, out PropertyRowSet_r? rows)
        {
            ErrorCodeValue result;
            byte[] auxIn = new byte[] { };
            ResolveNamesRequestBody resolveNamesRequestBody = new ResolveNamesRequestBody()
            {
                Reserved = reserved,
                HasState = true,
                State = stat,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            if (propTags != null)
            {
                LargePropTagArray propetyTags = new LargePropTagArray();
                propetyTags.PropertyTagCount = propTags.Value.CValues;
                propetyTags.PropertyTags = new PropertyTag[propetyTags.PropertyTagCount];
                for (int i = 0; i < propTags.Value.CValues; i++)
                {
                    propetyTags.PropertyTags[i].PropertyId = (ushort)((propTags.Value.AulPropTag[i] & 0xFFFF0000) >> 16);
                    propetyTags.PropertyTags[i].PropertyType = (ushort)(propTags.Value.AulPropTag[i] & 0x0000FFFF);
                }

                resolveNamesRequestBody.HasPropertyTags = true;
                resolveNamesRequestBody.PropertyTags = propetyTags;
            }
            else
            {
                resolveNamesRequestBody.HasPropertyTags = false;
                resolveNamesRequestBody.PropertyTags = new LargePropTagArray();
            }

            if (wstr != null)
            {
                resolveNamesRequestBody.HasNames = true;
                resolveNamesRequestBody.Names = wstr.Value;
            }
            else
            {
                resolveNamesRequestBody.HasNames = false;
            }

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(resolveNamesRequestBody, RequestType.ResolveNames);
            ResolveNamesResponseBody resolveNamesResponseBody = ResolveNamesResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)resolveNamesResponseBody.ErrorCode;
            if (resolveNamesResponseBody.RowCount != null)
            {
                PropertyRowSet_r newRows = AdapterHelper.ParsePropertyRowSet_r(resolveNamesResponseBody.PropertyTags.Value, resolveNamesResponseBody.RowCount.Value, resolveNamesResponseBody.RowData);
                rows = newRows;
            }
            else
            {
                rows = null;
            }

            if (resolveNamesResponseBody.HasMinimalIds)
            {
                PropertyTagArray_r propertyTagArray = new PropertyTagArray_r();
                propertyTagArray.CValues = resolveNamesResponseBody.MinimalIdCount.Value;
                propertyTagArray.AulPropTag = resolveNamesResponseBody.MinimalIds;
                mids = propertyTagArray;
            }
            else
            {
                mids = null;
            }

            return result;
        }

        /// <summary>
        /// The NspiGetTemplateInfo method returns information about template objects.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="type">A DWORD value. It specifies the display type of the template for which the information is requested.</param>
        /// <param name="dn">The value NULL or the DN of the template requested. The value is NULL-terminated.</param>
        /// <param name="codePage">A DWORD value. It specifies the code page of the template for which the information is requested.</param>
        /// <param name="localeID">A DWORD value. It specifies the LCID of the template for which the information is requested.</param>
        /// <param name="data">A reference to a PropertyRow_r value. On return, it contains the information requested.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue GetTemplateInfo(uint flags, uint type, string dn, uint codePage, uint localeID, out PropertyRow_r? data)
        {
            ErrorCodeValue result;
            byte[] auxIn = new byte[] { };
            GetTemplateInfoRequestBody getTemplateInfoRequestBody = new GetTemplateInfoRequestBody()
            {
                Flags = flags,
                DisplayType = type,
                CodePage = codePage,
                LocaleId = localeID,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            if (!string.IsNullOrEmpty(dn))
            {
                getTemplateInfoRequestBody.HasTemplateDn = true;
                getTemplateInfoRequestBody.TemplateDn = dn;
            }

            ChunkedResponse chunkedResponse = this.SendAddressBookRequest(getTemplateInfoRequestBody, RequestType.GetTemplateInfo);
            GetTemplateInfoResponseBody getTemplateInfoResponseBody = GetTemplateInfoResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
            result = (ErrorCodeValue)getTemplateInfoResponseBody.ErrorCode;
            if (getTemplateInfoResponseBody.HasRow)
            {
                PropertyRow_r propertyRow = AdapterHelper.ParsePropertyRow_r(getTemplateInfoResponseBody.Row.Value);
                data = propertyRow;
            }
            else
            {
                data = null;
            }

            return result;
        }
        #endregion

        #region Private method

        /// <summary>
        /// Send the request to address book server endpoint. 
        /// </summary>
        /// <param name="requestBody">The request body.</param>
        /// <param name="requestType">The type of the request.</param>
        /// <param name="cookieChange">Whether the session context cookie is changed.</param>
        /// <returns>The returned chunked response.</returns>
        private ChunkedResponse SendAddressBookRequest(IRequestBody requestBody, RequestType requestType, bool cookieChange = true)
        {
            byte[] rawBuffer = null;
            ChunkedResponse chunkedResponse = null;

            // Send the execute HTTP request and get the response
            HttpWebResponse response = MapiHttpAdapter.SendMAPIHttpRequest(this.site, this.addressBookUrl, this.userName, this.domainName, this.password, requestBody, requestType.ToString(), AdapterHelper.SessionContextCookies);
            rawBuffer = MapiHttpAdapter.ReadHttpResponse(response);
            string responseCode = response.Headers["X-ResponseCode"];
            this.site.Assert.AreEqual<uint>(0, uint.Parse(responseCode), "The request to the address book server should be executed successfully!");

            // Read the HTTP response buffer and parse the response to correct format
            chunkedResponse = ChunkedResponse.ParseChunkedResponse(rawBuffer);

            response.GetResponseStream().Close();
            if (cookieChange)
            {
                AdapterHelper.SessionContextCookies = response.Cookies;
            }

            return chunkedResponse;
        }

        /// <summary>
        /// Initialize Bind request body.
        /// </summary>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="flags">A set of bit flags that specify options to the server.</param>
        /// <returns>An instance of the Bind request body.</returns>
        private BindRequestBody BuildBindRequestBody(STAT stat, uint flags)
        {
            BindRequestBody bindRequestBody = new BindRequestBody();
            bindRequestBody.State = stat;
            bindRequestBody.Flags = flags;
            bindRequestBody.HasState = true;
            byte[] auxIn = new byte[] { };
            bindRequestBody.AuxiliaryBuffer = auxIn;
            bindRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            return bindRequestBody;
        }

        /// <summary>
        /// Initialize the Unbind request body.
        /// </summary>
        /// <returns>The Unbind request body</returns>
        private UnbindRequestBody BuildUnbindRequestBody()
        {
            UnbindRequestBody unbindRequest = new UnbindRequestBody();
            unbindRequest.Reserved = 0x00000000;
            byte[] auxIn = new byte[] { };
            unbindRequest.AuxiliaryBuffer = auxIn;
            unbindRequest.AuxiliaryBufferSize = (uint)auxIn.Length;

            return unbindRequest;
        }

        /// <summary>
        /// Build UpdateStat request body.
        /// </summary>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <returns>An instance of the UpdateStat request body.</returns>
        private UpdateStatRequestBody BuildUpdateStatRequestBody(STAT stat)
        {
            UpdateStatRequestBody updateStatRequestBody = new UpdateStatRequestBody();
            updateStatRequestBody.Reserved = 0x0;
            updateStatRequestBody.HasState = true;
            updateStatRequestBody.State = stat;
            updateStatRequestBody.DeltaRequested = true;

            byte[] auxIn = new byte[] { };
            updateStatRequestBody.AuxiliaryBuffer = auxIn;
            updateStatRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;
            return updateStatRequestBody;
        }

        /// <summary>
        /// Initialize the QueryColumns request body.
        /// </summary>
        /// <param name="flag">A set of bit flags that specify options to the server.</param>
        /// <returns>An instance of the QueryColumns request body.</returns>
        private QueryColumnsRequestBody BuildQueryColumnsRequestBody(uint flag)
        {
            QueryColumnsRequestBody queryColumnsRequestBody = new QueryColumnsRequestBody();
            queryColumnsRequestBody.MapiFlags = flag;
            queryColumnsRequestBody.Reserved = 0x0;
            byte[] auxIn = new byte[] { };
            queryColumnsRequestBody.AuxiliaryBuffer = auxIn;
            queryColumnsRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            return queryColumnsRequestBody;
        }

        /// <summary>
        /// Initialize the GetPropList request body.
        /// </summary>
        /// <param name="flag">A set of bit flags that specify options to the server.</param>
        /// <param name="mid">A minimal Entry ID structure that specifies the object for which to return properties.</param>
        /// <param name="codePage">An unsigned integer that specifies the code page that the server is being requested to use for string values of properties.</param>
        /// <returns>An instance of the GetPropList request body.</returns>
        private GetPropListRequestBody BuildGetPropListRequestBody(uint flag, uint mid, uint codePage)
        {
            GetPropListRequestBody getPropListRequestBody = new GetPropListRequestBody();
            getPropListRequestBody.Flags = flag;
            getPropListRequestBody.CodePage = codePage;
            getPropListRequestBody.MinmalId = mid;
            byte[] auxIn = new byte[] { };
            getPropListRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;
            getPropListRequestBody.AuxiliaryBuffer = auxIn;

            return getPropListRequestBody;
        }

        /// <summary>
        /// Initialize the GetProps request body.
        /// </summary>
        /// <param name="flags">A set of bit flags that specify options to the server.</param>
        /// <param name="hasState">A Boolean value that specifies whether the State field is present.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="hasPropertyTags">A Boolean value that specifies whether the PropertyTags field is present.</param>
        /// <param name="propetyTags">A LargePropertyTagArray structure that contains the property tags of the properties that the client is requesting.</param>
        /// <returns>An instance of the GetProps request body.</returns>
        private GetPropsRequestBody BuildGetPropsRequestBody(uint flags, bool hasState, STAT? stat, bool hasPropertyTags, LargePropTagArray propetyTags)
        {
            GetPropsRequestBody getPropertyRequestBody = new GetPropsRequestBody();
            getPropertyRequestBody.Flags = flags;
            byte[] auxIn = new byte[] { };
            getPropertyRequestBody.AuxiliaryBuffer = auxIn;
            getPropertyRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            if (hasState)
            {
                getPropertyRequestBody.HasState = true;
                if (stat != null)
                {
                    getPropertyRequestBody.State = (STAT)stat;
                }
            }
            else
            {
                getPropertyRequestBody.HasState = false;
                STAT statNew = new STAT();
                getPropertyRequestBody.State = statNew;
            }

            if (hasPropertyTags)
            {
                getPropertyRequestBody.HasPropertyTags = true;
                getPropertyRequestBody.PropertyTags = propetyTags;
            }
            else
            {
                getPropertyRequestBody.HasPropertyTags = false;
                LargePropTagArray propetyTagsNew = new LargePropTagArray();
                getPropertyRequestBody.PropertyTags = propetyTagsNew;
            }

            return getPropertyRequestBody;
        }
        #endregion
    }
}