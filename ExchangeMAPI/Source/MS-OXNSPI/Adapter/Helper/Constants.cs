namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    /// <summary>
    /// Constants used in MS-OXNSPI.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Size of FlatUID_r structure in byte.
        /// </summary>
        public const int FlatUIDByteSize = 16;

        /// <summary>
        /// The maximum number of rows for the NspiGetMatches method to return in a restricted address book container.
        /// </summary>
        public const uint GetMatchesRequestedRowNumber = 5000;

        /// <summary>
        /// The maximum number of rows for the NspiQueryRows method to return in a restricted address book container.
        /// </summary>
        public const uint QueryRowsRequestedRowNumber = 5000;

        /// <summary>
        /// A string which specifies a user name which doesn't exist.
        /// </summary>
        public const string UnresolvedName = "XXXXXX";

        /// <summary>
        /// A CodePage that server does not recognize (0xFFFFFFFF).
        /// </summary>
        public const string UnrecognizedCodePage = "4294967295";

        /// <summary>
        /// A Minimal Entry ID that server does not recognize (0xFFFFFFFF).
        /// </summary>
        public const string UnrecognizedMID = "4294967295";
    }
}