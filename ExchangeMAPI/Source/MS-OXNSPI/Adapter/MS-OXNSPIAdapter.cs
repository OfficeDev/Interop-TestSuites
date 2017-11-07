namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Net;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter implementation. This class implements the methods defined in the interface IMS_OXNSPIAdapter. 
    /// </summary>
    public partial class MS_OXNSPIAdapter : ManagedAdapterBase, IMS_OXNSPIAdapter
    {
        #region Variables

        /// <summary>
        /// The RPC binding.
        /// </summary>
        private IntPtr rpcBinding = IntPtr.Zero;

        /// <summary>
        /// The RPC context handle.
        /// </summary>
        private IntPtr contextHandle = IntPtr.Zero;

        /// <summary>
        /// Original server name
        /// </summary>
        private string originalServerName;

        /// <summary>
        /// The Mailbox userName which can be used by client to connect to the SUT.
        /// </summary>
        private string userName;

        /// <summary>
        /// The user password which can be used by client to connect to the SUT.
        /// </summary>
        private string password;

        /// <summary>
        /// Define the name of domain where the server belongs to.
        /// </summary>
        private string domainName;

        /// <summary>
        /// Public structure which contains auto discover properties.
        /// </summary>
        private AutoDiscoverProperties autoDiscoverProperties;

        /// <summary>
        /// The NspiRpcAdapter contains the RPC implements for the interfaces of IMS_OXNSPIAdapter.
        /// </summary>
        private NspiRpcAdapter nspiRpcAdapter;

        /// <summary>
        /// The MapiHttpAdapter contains the MAPIHTTP implements for the interfaces of IMS_OXNSPIAdapter.
        /// </summary>
        private NspiMapiHttpAdapter nspiMapiHttpAdapter;

        /// <summary>
        /// The transport used by the test suite.
        /// </summary>
        private string transport;

        /// <summary>
        /// The time internal that is used to wait to retry when the returned error code is GeneralFailure.
        /// </summary>
        private int waitTime;

        /// <summary>
        /// The retry count that is used to retry when the returned error code is GeneralFailure.
        /// </summary>
        private uint maxRetryCount;

        #endregion

        #region Override functions
        /// <summary>
        /// Initializes the current adapter instance associated with a test site.
        /// </summary>
        /// <param name="testSite">The test site instance associated with the current adapter.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXNSPI";
            Common.MergeConfiguration(testSite);
            AdapterHelper.Site = testSite;
            if (bool.Parse(Common.GetConfigurationPropertyValue("MS-OXNSPI_Supported", this.Site)))
            {
                this.nspiRpcAdapter = null;
                this.nspiMapiHttpAdapter = null;
                this.originalServerName = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
                this.userName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
                this.Site.Assume.IsNotNull(this.userName, "The User1Name field in the ptfconfig file should not be null.");
                this.Site.Assume.IsNotNull(Common.GetConfigurationPropertyValue("User2Name", this.Site), "The User2Name field in the ptfconfig file should not be null.");
                this.Site.Assume.IsNotNull(Common.GetConfigurationPropertyValue("User3Name", this.Site), "The User3Name field in the ptfconfig file should not be null.");
                this.domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
                this.password = Common.GetConfigurationPropertyValue("User1Password", this.Site);
                string publicFolderMailbox = Common.GetConfigurationPropertyValue("PublicFolderMailbox", this.Site);
                string requestURL = Common.GetConfigurationPropertyValue("AutoDiscoverUrlFormat", this.Site);
                requestURL = Regex.Replace(requestURL, @"\[ServerName\]", this.originalServerName, RegexOptions.IgnoreCase);
                this.transport = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower(System.Globalization.CultureInfo.CurrentCulture);
                this.Site.Assert.IsTrue(this.transport == "mapi_http" || this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp", @"The TransportSeq field in the ptfconfig file must be set to one of the following three values: mapi_http, ncacn_http and ncacn_ip_tcp.");
                this.waitTime = Convert.ToInt32(Common.GetConfigurationPropertyValue("WaitTime", testSite));
                this.maxRetryCount = Convert.ToUInt32(Common.GetConfigurationPropertyValue("RetryCount", testSite));
                AdapterHelper.Transport = this.transport;
                if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
                {
                    this.InitializeRPC();
                    this.nspiRpcAdapter = new NspiRpcAdapter(this.Site, this.rpcBinding, this.contextHandle, this.waitTime, this.maxRetryCount);
                }
                else
                {
                    AdapterHelper.SessionContextCookies = new CookieCollection();
                    this.autoDiscoverProperties = AutoDiscover.GetAutoDiscoverProperties(this.Site, this.originalServerName, this.userName, this.domainName, requestURL, this.transport, publicFolderMailbox);
                    this.Site.Assert.IsNotNull(this.autoDiscoverProperties.AddressBookUrl, @"The auto discover process should return the URL to be used to connect with a NSPI server through MAPI over HTTP successfully.");
                    this.nspiMapiHttpAdapter = new NspiMapiHttpAdapter(this.Site, this.userName, this.password, this.domainName, this.autoDiscoverProperties.AddressBookUrl);
                }
            }
        }

        /// <summary>
        /// Release the test site.
        /// </summary>
        public override void Reset()
        {
            // Destroy the context handle if it is not destroyed in the test case.
            if ((this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp") && this.contextHandle != IntPtr.Zero)
            {
                this.NspiUnbind(0);
            }

            if (AdapterHelper.SessionContextCookies.Count != 0)
            {
                AdapterHelper.SessionContextCookies = new CookieCollection();
            }

            base.Reset();
        }

        #endregion

        #region Instance interface

        /// <summary>
        /// The NspiBind method initiates a session between a client and the server.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="serverGuid">The value NULL or a pointer to a GUID value that is associated with the specific server.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiBind(uint flags, STAT stat, ref FlatUID_r? serverGuid, bool needRetry = true)
        {
            ErrorCodeValue result;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiBind(flags, stat, ref serverGuid, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.Bind(flags, stat, ref serverGuid);
            }

            this.VerifyNspiBind();
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiUnbind method destroys the context handle. No other action is taken.
        /// </summary>
        /// <param name="reserved">A DWORD [MS-DTYP] value reserved for future use. This property is ignored by the server.</param>
        /// <returns>A DWORD value that specifies the return status of the method.</returns>
        public uint NspiUnbind(uint reserved)
        {
            uint result;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiUnbind(reserved, ref this.contextHandle);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.Unbind(reserved);
            }

            this.VerifyNspiUnbind(this.contextHandle);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiGetSpecialTable method returns the rows of a special table to the client. 
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="version">A reference to a DWORD. On input, it holds the value of the version number of
        /// the address book hierarchy table that the client has. On output, it holds the version of the server's address book hierarchy table.</param>
        /// <param name="rows">A PropertyRowSet_r structure. On return, it holds the rows for the table that the client is requesting.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetSpecialTable(uint flags, ref STAT stat, ref uint version, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            ErrorCodeValue result;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiGetSpecialTable(flags, ref stat, ref version, out rows, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.GetSpecialTable(flags, ref stat, ref version, out rows);
            }

            this.VerifyNspiGetSpecialTable(result, rows);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiUpdateStat method updates the STAT block that represents the position in a table 
        /// to reflect positioning changes requested by the client.
        /// </summary>
        /// <param name="reserved">A DWORD value. Reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A pointer to a STAT block describing a logical position in a specific address book container.</param>
        /// <param name="delta">The value NULL or a pointer to a LONG value that indicates movement 
        /// within the address book container specified by the input parameter stat.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiUpdateStat(uint reserved, ref STAT stat, ref int? delta, bool needRetry = true)
        {
            ErrorCodeValue result;
            STAT inputStat = stat;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiUpdateStat(reserved, ref stat, ref delta, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.UpdateStat(ref stat, ref delta);
            }

            this.VerifyNspiUpdateStat(result, inputStat, stat);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiQueryColumns method returns a list of all the properties that the server is aware of. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="columns">A PropertyTagArray_r structure that contains a list of proptags.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiQueryColumns(uint reserved, uint flags, out PropertyTagArray_r? columns, bool needRetry = true)
        {
            ErrorCodeValue result;

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiQueryColumns(reserved, flags, out columns, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.QueryColumns(flags, out columns);
            }

            this.VerifyNspiQueryColumns(result, columns);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiGetPropList method returns a list of all the properties that have values on a specified object.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="mid">A DWORD value that contains a Minimal Entry ID.</param>
        /// <param name="codePage">The code page in which the client wants the server to express string values properties.</param>
        /// <param name="propTags">A PropertyTagArray_r value. On return, it holds a list of properties.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetPropList(uint flags, uint mid, uint codePage, out PropertyTagArray_r? propTags, bool needRetry = true)
        {
            ErrorCodeValue result;

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiGetPropList(flags, mid, codePage, out propTags, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.GetPropList(flags, mid, codePage, out propTags);
            }

            this.VerifyNspiGetPropList(result, propTags);
            this.VerifyTransport();
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
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetProps(uint flags, STAT stat, PropertyTagArray_r? propTags, out PropertyRow_r? rows, bool needRetry = true)
        {
            ErrorCodeValue result;

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiGetProps(flags, stat, propTags, out rows, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.GetProps(flags, stat, propTags, out rows);
            }

            this.VerifyNspiGetProps(result, rows);
            this.VerifyTransport();
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
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiQueryRows(uint flags, ref STAT stat, uint tableCount, uint[] table, uint count, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            ErrorCodeValue result;
            STAT inputStat = stat;

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiQueryRows(flags, ref stat, tableCount, table, count, propTags, out rows, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.QueryRows(flags, ref stat, tableCount, table, count, propTags, out rows);
            }

            this.VerifyNspiQueryRows(result, rows, inputStat, stat);
            this.VerifyTransport();
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
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiSeekEntries(uint reserved, ref STAT stat, PropertyValue_r target, PropertyTagArray_r? table, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            ErrorCodeValue result = 0;
            STAT inputStat = stat;

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiSeekEntries(reserved, ref stat, target, table, propTags, out rows, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.SeekEntries(reserved, ref stat, target, table, propTags, out rows);
            }

            this.VerifyNspiSeekEntries(result, rows, inputStat, stat);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiGetMatches method returns an Explicit Table. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use.</param>
        /// <param name="stat">A STAT block describing a logical position in a specific address book container.</param>
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
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetMatches(uint reserved, ref STAT stat, PropertyTagArray_r? proReserved, uint reserved2, Restriction_r? filter, PropertyName_r? propName, uint requested, out PropertyTagArray_r? outMids, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            ErrorCodeValue result = 0;
            STAT inputStat = stat;

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiGetMatches(reserved, ref stat, proReserved, reserved2, filter, propName, requested, out outMids, propTags, out rows, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.GetMatches(reserved, ref stat, proReserved, reserved2, filter, propName, requested, out outMids, propTags, out rows);
            }

            this.VerifyNspiGetMatches(result, rows, outMids, inputStat, stat);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiResortRestriction method applies to a sort order to the objects in a restricted address book container.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A reference to a STAT block describing a logical position in a specific address book container.</param>
        /// <param name="proInMIds">A PropertyTagArray_r value. 
        /// It holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="outMIds">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs 
        /// that comprise a restricted address book container.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiResortRestriction(uint reserved, ref STAT stat, PropertyTagArray_r proInMIds, ref PropertyTagArray_r? outMIds, bool needRetry = true)
        {
            ErrorCodeValue result;
            STAT inputStat = stat;

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiResortRestriction(reserved, ref stat, proInMIds, ref outMIds, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.ResortRestriction(reserved, ref stat, proInMIds, ref outMIds);
            }

            this.VerifyNspiResortRestriction(result, outMIds, inputStat, stat);
            this.VerifyTransport();
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
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiCompareMIds(uint reserved, STAT stat, uint mid1, uint mid2, out int results, bool needRetry = true)
        {
            ErrorCodeValue result;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiCompareMIds(reserved, stat, mid1, mid2, out results, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.CompareMIds(reserved, stat, mid1, mid2, out results);
            }

            this.VerifyNspiCompareMIds(result);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiDNToMId method maps a set of DN to a set of Minimal Entry ID.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="names">A StringsArray_r value. It holds a list of strings that contain DNs.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiDNToMId(uint reserved, StringsArray_r names, out PropertyTagArray_r? mids, bool needRetry = true)
        {
            ErrorCodeValue result;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiDNToMId(reserved, names, out mids, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.DNToMId(reserved, names, out mids);
            }

            this.VerifyNspiDNToMId(result, mids);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiModProps method is used to modify the properties of an object in the address book. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r. 
        /// It contains a list of the proptags of the columns from which the client requests all the values to be removed.</param>
        /// <param name="row">A PropertyRow_r value. It contains an address book row.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiModProps(uint reserved, STAT stat, PropertyTagArray_r? propTags, PropertyRow_r row, bool needRetry = true)
        {
            ErrorCodeValue result;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiModProps(reserved, stat, propTags, row, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.ModProps(stat, propTags, row);
            }

            this.VerifyNspiModProps(result);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiModLinkAtt method modifies the values of a specific property of a specific row in the address book.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="propTag">A DWORD value. It contains the proptag of the property that the client wants to modify.</param>
        /// <param name="mid">A DWORD value that contains the Minimal Entry ID of the address book row that the client wants to modify.</param>
        /// <param name="entryIds">A BinaryArray value. It contains a list of EntryIDs to be used to modify the requested property on the requested address book row.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiModLinkAtt(uint flags, uint propTag, uint mid, BinaryArray_r entryIds, bool needRetry = true)
        {
            ErrorCodeValue result;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiModLinkAtt(flags, propTag, mid, entryIds, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.ModLinkAtt(flags, propTag, mid, entryIds);
            }

            this.VerifyNspiModLinkAtt(result);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiResolveNames method takes a set of string values in an 8-bit character set and performs ANR on those strings. 
        /// The NspiResolveNames method taking string values in an 8-bit character set is not supported when mapi_http transport is used. 
        /// </summary>
        /// <param name="reserved">A DWORD reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value containing a list of the proptags of the columns 
        /// that the client requests to be returned for each row returned.</param>
        /// <param name="stringArray">A StringsArray_r value. It specifies the values on which the client is requesting the server to do ANR.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it contains a list of Minimal Entry IDs that match the array of strings.</param>
        /// <param name="rows">A reference to a PropertyRowSet_r value. 
        /// It contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiResolveNames(uint reserved, STAT stat, PropertyTagArray_r? propTags, StringsArray_r? stringArray, out PropertyTagArray_r? mids, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            ErrorCodeValue result;
            result = this.nspiRpcAdapter.NspiResolveNames(reserved, stat, propTags, stringArray, out mids, out rows, needRetry);

            this.VerifyNspiResolveNames(result, mids, rows);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// The NspiResolveNamesW method takes a set of string values in the Unicode character set and performs ANR on those strings. 
        /// </summary>
        /// <param name="reserved">A DWORD value that is reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r containing a list of the proptags of the columns 
        /// that the client requests to be returned for each row returned.</param>
        /// <param name="wstr">A WStringsArray_r value. It specifies the values on which the client is requesting the server to perform ANR.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it contains a list of Minimal Entry IDs that match the array of strings.</param>
        /// <param name="rowOfResolveNamesW">A reference to a PropertyRowSet_r structure. It contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiResolveNamesW(uint reserved, STAT stat, PropertyTagArray_r? propTags, WStringsArray_r? wstr, out PropertyTagArray_r? mids, out PropertyRowSet_r? rowOfResolveNamesW, bool needRetry = true)
        {
            ErrorCodeValue result;

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiResolveNamesW(reserved, stat, propTags, wstr, out mids, out rowOfResolveNamesW, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.ResolveNames(reserved, stat, propTags, wstr, out mids, out rowOfResolveNamesW);
            }

            this.VerifyNspiResolveNamesW(result, mids, rowOfResolveNamesW);
            this.VerifyTransport();
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
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetTemplateInfo(uint flags, uint type, string dn, uint codePage, uint localeID, out PropertyRow_r? data, bool needRetry = true)
        {
            ErrorCodeValue result;

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiGetTemplateInfo(flags, type, dn, codePage, localeID, out data, needRetry);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.GetTemplateInfo(flags, type, dn, codePage, localeID, out data);
            }

            this.VerifyNspiGetTemplateInfo(result, flags, data);
            this.VerifyTransport();
            return result;
        }
        #endregion

        #region IDisposable Members
        /// <summary>
        /// This method is used to implement clean-up codes.
        /// </summary>
        /// <param name="disposing">Set TRUE to dispose resource otherwise set FALSE.</param>
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            OxnspiInterop.RpcBindingFree(ref this.rpcBinding);
        }
        #endregion

        #region Initialize RPC.
        /// <summary>
        /// Initialize the client and server and build the transport tunnel between client and server.
        /// </summary>
        private void InitializeRPC()
        {
            string serverName = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
            string userName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            string domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
            string password = Common.GetConfigurationPropertyValue("User1Password", this.Site);

            // Create identity for the user to connect to the server.
            OxnspiInterop.CreateIdentity(
                domainName,
                userName,
                password);
            MapiContext rpcContext = MapiContext.GetDefaultRpcContext(this.Site);

            // Create Service Principal Name (SPN) string for the user to connect to the server.
            string userSpn = string.Empty;
            userSpn = Regex.Replace(rpcContext.SpnFormat, @"\[ServerName\]", serverName, RegexOptions.IgnoreCase);

            // Bind the client to RPC server.
            uint status = OxnspiInterop.BindToServer(serverName, rpcContext.AuthenLevel, rpcContext.AuthenService, rpcContext.TransportSequence, rpcContext.RpchUseSsl, rpcContext.RpchAuthScheme, userSpn, null, rpcContext.SetUuid);
            this.Site.Assert.AreEqual<uint>(0, status, "Create binding handle with server {0} should success!", serverName);
            this.rpcBinding = OxnspiInterop.GetBindHandle();
            this.Site.Assert.AreNotEqual<IntPtr>(IntPtr.Zero, this.rpcBinding, "A valid RPC Binding handle is needed!");
        }
        #endregion
    }
}