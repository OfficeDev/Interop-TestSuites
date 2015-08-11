namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    /// <summary>
    /// Used to specify the transport type for soap messages.
    /// </summary>
    public enum TransportType
    {
        /// <summary>
        /// Indicate the soap transport over http.
        /// </summary>
        HTTP,

        /// <summary>
        /// Indicate the soap transport over https.
        /// </summary>
        HTTPS
    }

    /// <summary>
    /// This enum indicate the option of authentication information client used.
    /// </summary>
    public enum UserAuthenticationOption
    {
        /// <summary>
        /// Specify Adapter use an authenticated account.
        /// </summary>
        Authenticated,

        /// <summary>
        /// Specify Adapter use an unauthenticated account.
        /// </summary>
        Unauthenticated
    }
}