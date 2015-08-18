namespace Microsoft.Protocols.TestSuites.Common
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MAPI CALL Context
    /// </summary>
    public class MapiContext
    {
        /// <summary>
        /// Gets or sets the instance of ITestSite
        /// </summary>
        public ITestSite TestSite { get; set; }

        /// <summary>
        /// Gets or sets the authentication level for creating RPC binding
        /// </summary>
        public uint AuthenLevel { get; set; }

        /// <summary>
        /// Gets or sets the authentication services by identifying the security package that provides the service
        /// </summary>
        public uint AuthenService { get; set; }

        /// <summary>
        /// Gets or sets RPC transport sequence type
        /// </summary>
        public string TransportSequence { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to use RPC over HTTP with SSL, true to use RPC over HTTP with SSL, false to use RPC over HTTP without SSL.
        /// </summary>
        public bool RpchUseSsl { get; set; }

        /// <summary>
        /// Gets or sets the authentication scheme used in the http authentication for RPC over HTTP.
        /// </summary>
        public string RpchAuthScheme { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether use obfuscation.
        /// </summary>
        public bool Xor { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether use compress.
        /// </summary>
        public bool Compress { get; set; }

        /// <summary>
        /// Gets or sets the Service Principal Name (SPN) format
        /// </summary>
        public string SpnFormat { get; set; }

        /// <summary>
        /// Gets or sets exchange server version
        /// </summary>
        public ushort[] EXServerVersion { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to set uuid flag, true to set PFC_OBJECT_UUID(0x80) field of RPC header, false to not set this field
        /// </summary>
        public bool SetUuid { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to enable auto redirect, true indicates enable auto redirect, false indicates disable auto redirect
        /// </summary>
        public bool AutoRedirect { get; set; }

        /// <summary>
        /// Gets or sets the code page in which text data is sent if the Unicode format is not requested by the client on subsequent calls that use this Session Context. This will be used in ulCpid parameter of EcDoConnectEx method, as defined in [MS-OXCRPC] section 3.1.4.11.
        /// </summary>
        public uint? CodePageId { get; set; }

        /// <summary>
        /// Get default RPC CALL Context
        /// </summary>
        /// <param name="site">The instance of ITestSite</param>
        /// <returns>RPC CALL Context</returns>
        public static MapiContext GetDefaultRpcContext(ITestSite site)
        {
            MapiContext mapiContext = new MapiContext
            {
                AuthenLevel = uint.Parse(Common.GetConfigurationPropertyValue("RpcAuthenticationLevel", site)),
                AuthenService = uint.Parse(Common.GetConfigurationPropertyValue("RpcAuthenticationService", site)),
                TransportSequence = Common.GetConfigurationPropertyValue("TransportSeq", site),
                SpnFormat = Common.GetConfigurationPropertyValue("ServiceSpnFormat", site)
            };
            if (mapiContext.TransportSequence.ToLower() == "ncacn_http")
            {
                bool rpchUseSsl;
                if (!bool.TryParse(Common.GetConfigurationPropertyValue("RpchUseSsl", site), out rpchUseSsl))
                {
                    site.Assert.Fail("Value of 'RpchUseSsl' property is invalid.");
                }

                mapiContext.RpchUseSsl = rpchUseSsl;
                mapiContext.RpchAuthScheme = Common.GetConfigurationPropertyValue("RpchAuthScheme", site);
                if (mapiContext.RpchAuthScheme.ToLower() != "basic" && mapiContext.RpchAuthScheme.ToLower() != "ntlm")
                {
                    site.Assert.Fail("Value of 'RpchAuthScheme' property is invalid.");
                }
            }

            mapiContext.SetUuid = bool.Parse(Common.GetConfigurationPropertyValue("SetUuid", site));
            mapiContext.TestSite = site;
            mapiContext.EXServerVersion = new ushort[3] { 0, 0, 0 };
            mapiContext.AutoRedirect = true;
            mapiContext.CodePageId = null;

            return mapiContext;
        }
    }
}