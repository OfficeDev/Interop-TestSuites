namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    /// <summary>
    /// This enumeration specifies the type of version.
    /// </summary>
    public enum VersionType
    {
        /// <summary>
        /// Check in minor version.
        /// </summary>
        MinorCheckIn,
        
        /// <summary>
        /// Check in major version.
        /// </summary>
        MajorCheckIn,

        /// <summary>
        /// Check in as overwrite.
        /// </summary>
        OverwriteCheckIn,
    }

    /// <summary>
    /// This enumeration indicates the protocol operation names.
    /// </summary>
    public enum OperationName
    {
        /// <summary>
        /// The DeleteAllVersions operation is used to delete all the previous versions of the specified file.
        /// </summary>
        DeleteAllVersions,

        /// <summary>
        /// The DeleteVersion operation is used to delete a specific version of the specified file.
        /// </summary>
        DeleteVersion,

        /// <summary>
        /// The GetVersions operation is used to get details about
        /// all versions of the specified file that the user can access.
        /// </summary>
        GetVersions,

        /// <summary>
        /// The RestoreVersion operation is used to restore the specified file to a specific version.
        /// </summary>
        RestoreVersion,
    }

    /// <summary>
    /// This enumeration indicates the client uses which protocol to transport data.
    /// </summary>
    public enum TransportProtocol
    {
        /// <summary>
        /// Specify that the client uses HTTP to transport data.
        /// </summary>
        HTTP,

        /// <summary>
        /// Specify that the client uses HTTPS to transport data.
        /// </summary>
        HTTPS
    }
}