namespace Microsoft.Protocols.TestSuites.MS_OXORULE
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
                propetyTags.PropertyTagCount = propTags.Value.Values;
                propetyTags.PropertyTags = new PropertyTag[propetyTags.PropertyTagCount];
                for (int i = 0; i < propTags.Value.Values; i++)
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
        #endregion
    }
}