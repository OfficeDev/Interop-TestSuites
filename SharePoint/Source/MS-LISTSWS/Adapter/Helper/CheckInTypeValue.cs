namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{   
    /// <summary>
    /// The values of various check in types.
    /// </summary>
    public struct CheckInTypeValue
    {
        /// <summary>
        /// A string indicates the value which means Minor check in
        /// </summary>
        public const string MinorCheckIn = "0";

        /// <summary>
        /// A string indicates the value which means Major check in
        /// </summary>
        public const string MajorCheckIn = "1";

        /// <summary>
        /// A string indicates the value which means Overwrite check in
        /// </summary>
        public const string OverwriteCheckIn = "2";

        /// <summary>
        /// A string indicates the value which is invalid for Check in Type value in protocol SUT
        /// </summary>
        public const string InvalidValue = "-1";
    }
}