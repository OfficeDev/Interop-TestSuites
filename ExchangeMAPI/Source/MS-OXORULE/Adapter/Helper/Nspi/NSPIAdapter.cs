namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Net;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter implementation. This class implements the methods defined in the interface IMS_OXNSPIAdapter. 
    /// </summary>
    public class NSPIAdapter
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
        private int waitTime = 0;

        /// <summary>
        /// The retry count that is used to retry when the returned error code is GeneralFailure.
        /// </summary>
        private uint maxRetryCount = 0;

        private ITestSite site;
        #endregion

        public NSPIAdapter(ITestSite site)
        {
            this.site = site;
            this.originalServerName = Common.GetConfigurationPropertyValue("SutComputerName", this.site);
            this.userName = Common.GetConfigurationPropertyValue("AdminUserName", this.site);
            this.domainName = Common.GetConfigurationPropertyValue("Domain", this.site);
            this.password = Common.GetConfigurationPropertyValue("AdminUserPassword", this.site);
            this.waitTime= int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.site));
            this.maxRetryCount = uint.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.site));
            string requestURL = Common.GetConfigurationPropertyValue("AutoDiscoverUrlFormat", this.site);
            requestURL = Regex.Replace(requestURL, @"\[ServerName\]", this.originalServerName, RegexOptions.IgnoreCase);
            this.transport = Common.GetConfigurationPropertyValue("TransportSeq", this.site).ToLower(System.Globalization.CultureInfo.CurrentCulture);
            AdapterHelper.Transport = this.transport;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                this.InitializeRPC();
                this.nspiRpcAdapter = new NspiRpcAdapter(this.site, this.rpcBinding, this.contextHandle, this.waitTime, this.maxRetryCount);
            }
            else
            {
                AdapterHelper.SessionContextCookies = new CookieCollection();
                this.autoDiscoverProperties = AutoDiscover.GetAutoDiscoverProperties(this.site, this.originalServerName, this.userName, this.domainName, requestURL, this.transport);
                this.site.Assert.IsNotNull(this.autoDiscoverProperties.AddressBookUrl, @"The auto discover process should return the URL to be used to connect with a NSPI server through MAPI over HTTP successfully.");
                this.nspiMapiHttpAdapter = new NspiMapiHttpAdapter(this.site, this.userName, this.password, this.domainName, this.autoDiscoverProperties.AddressBookUrl);
            }
        }

        /// <summary>
        /// The NspiBind method initiates a session between a client and the server.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="serverGuid">The value NULL or a pointer to a GUID value that is associated with the specific server.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiBind(uint flags, STAT stat, ref FlatUID_r? serverGuid)
        {
            ErrorCodeValue result;
            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                result = this.nspiRpcAdapter.NspiBind(flags, stat, ref serverGuid);
            }
            else
            {
                result = this.nspiMapiHttpAdapter.Bind(flags, stat, ref serverGuid);
            }

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

            return result;
        }

        #region Initialize RPC.
        /// <summary>
        /// Initialize the client and server and build the transport tunnel between client and server.
        /// </summary>
        private void InitializeRPC()
        {
            string serverName = Common.GetConfigurationPropertyValue("SutComputerName", this.site);
            string userName = Common.GetConfigurationPropertyValue("AdminUserName", this.site);
            string domainName = Common.GetConfigurationPropertyValue("Domain", this.site);
            string password = Common.GetConfigurationPropertyValue("AdminUserPassword", this.site);

            // Create identity for the user to connect to the server.
            OxnspiInterop.CreateIdentity(
                domainName,
                userName,
                password);
            MapiContext rpcContext = MapiContext.GetDefaultRpcContext(this.site);

            // Create Service Principal Name (SPN) string for the user to connect to the server.
            string userSpn = string.Empty;
            userSpn = Regex.Replace(rpcContext.SpnFormat, @"\[ServerName\]", serverName, RegexOptions.IgnoreCase);

            // Bind the client to RPC server.
            uint status = OxnspiInterop.BindToServer(serverName, rpcContext.AuthenLevel, rpcContext.AuthenService, rpcContext.TransportSequence, rpcContext.RpchUseSsl, rpcContext.RpchAuthScheme, userSpn, null, rpcContext.SetUuid);
            this.site.Assert.AreEqual<uint>(0, status, "Create binding handle with server {0} should success!", serverName);
            this.rpcBinding = OxnspiInterop.GetBindHandle();
            this.site.Assert.AreNotEqual<IntPtr>(IntPtr.Zero, this.rpcBinding, "A valid RPC Binding handle is needed!");
        }
        #endregion
    }
}