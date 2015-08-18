namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// Error code returned by RPC calls.
    /// </summary>
    public static class OxcRpcErrorCode
    {
        /// <summary>
        /// The operation succeeded.
        /// </summary>
        public const uint ECNone = 0x00000000;

        /// <summary>
        /// A badly formatted RPC buffer was detected.
        /// </summary>
        public const uint ECRpcFormat = 0x000004B6;

        /// <summary>
        /// (0x0000047D)Error code returned when response payload bigger than the maximum of pcbOut.
        /// </summary>
        public const uint ECResponseTooBig = 0x0000047D;

        /// <summary>
        /// (0x0000047D)Error code returned when a buffer passed to this function is not big enough.
        /// </summary>
        public const uint ECBufferTooSmall = 0x0000047D;
    }
}