namespace Microsoft.Protocols.TestSuites.MS_OXCNOTIF
{
    /// <summary>
    /// Address families.
    /// </summary>
    public enum AddressFamily : ushort
    {
        /// <summary>
        /// Internetwork: UDP, TCP, etc.
        /// </summary>
        AF_INET = 2,

        /// <summary>
        /// Internetwork Version 6
        /// </summary>
        AF_INET6 = 23,

        /// <summary>
        /// Invalid address family
        /// </summary>
        Invalid = 0xff
    }
}