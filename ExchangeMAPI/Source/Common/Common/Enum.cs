namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// Specify how data that follows this header MUST be interpreted.
    /// </summary>
    [System.Flags]
    public enum RpcHeaderExtFlags : short
    {
        /// <summary>
        /// All flags are not set.
        /// </summary>
        None = 0,

        /// <summary>
        /// The data that follows the RPC_HEADER_EXT is compressed.
        /// </summary>
        Compressed = 0x0001,

        /// <summary>
        /// The data following the RPC_HEADER_EXT has been obfuscated.
        /// </summary>
        XorMagic = 0x0002,

        /// <summary>
        /// Indicates that no other RPC_HEADER_EXT follows the data of the current RPC_HEADER_EXT.
        /// </summary>
        Last = 0x0004
    }

    /// <summary>
    /// The logon type.
    /// </summary>
    public enum LogonType
    {
        /// <summary>
        /// Mailbox type.
        /// </summary>
        Mailbox,

        /// <summary>
        /// Public folder type.
        /// </summary>
        PublicFolder
    }

    /// <summary>
    /// The type of connection
    /// </summary>
    public enum ConnectionType
    {
        /// <summary>
        /// Connect to server for private mailbox
        /// </summary>
        PrivateMailboxServer = 1,

        /// <summary>
        /// Connect to server for public folder
        /// </summary>
        PublicFolderServer = 0
    }
}