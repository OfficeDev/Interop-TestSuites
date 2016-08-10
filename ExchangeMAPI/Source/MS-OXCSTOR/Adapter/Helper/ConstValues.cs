namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    /// <summary>
    /// A Utility class intends to provide some useful methods with capture
    /// and also contains some important variable to retrieve the deployment data
    /// So that can achieve easy-maintainability and easy-expandability
    /// </summary>
    public class ConstValues
    {
        #region Const Global variables

        /// <summary>
        /// The SUT domain
        /// </summary>
        public const string Domain = "Domain";

        /// <summary>
        /// The user name
        /// </summary>
        public const string UserName = "AdminUserName";

        /// <summary>
        /// The user name on Server1
        /// </summary>
        public const string UserNameOfMailboxOnServer1 = "UserNameOfMailboxOnServer1";

        /// <summary>
        /// The user password on Server1
        /// </summary>
        public const string UserPasswordOfMailboxOnServer1 = "UserPasswordOfMailboxOnServer1";

        /// <summary>
        /// The user name on Server2
        /// </summary>
        public const string UserNameOfMailboxOnServer2 = "UserNameOfMailboxOnServer2";

        /// <summary>
        /// The user password on Server2
        /// </summary>
        public const string UserPasswordOfMailboxOnServer2 = "UserPasswordOfMailboxOnServer2";

        /// <summary>
        /// The user's ESSDN
        /// </summary>
        public const string UserEssdn = "UserEssdn";

        /// <summary>
        /// The test user1's ESSDN
        /// </summary>
        public const string User1ESSDN = "UserEssdnOfMailboxOnServer1";

        /// <summary>
        /// The test user2's ESSDN
        /// </summary>
        public const string User2ESSDN = "UserEssdnOfMailboxOnServer2";

        /// <summary>
        /// The SUT user password
        /// </summary>
        public const string Password = "UserPassword";

        /// <summary>
        /// The name of server1
        /// </summary>
        public const string Server1 = "SutComputerName";

        /// <summary>
        /// The name of server2
        /// </summary>
        public const string Server2 = "Sut2ComputerName";

        /// <summary>
        /// The transport protocol sequence
        /// </summary>
        public const string TransportSeq = "TransportSeq";

        /// <summary>
        /// The public database name
        /// </summary>
        public const string PublicDbNameOnServer1 = "PublicDbNameOnServer1";

        /// <summary>
        /// The user whose mailbox will be disabled.
        /// </summary>
        public const string UserForDisableMailbox = "UserNameForDisableMailbox";

        /// <summary>
        /// The password of the user whose mailbox will be disabled.
        /// </summary>
        public const string PasswordForDisableMailbox = "UserPasswordForDisableMailbox";

        /// <summary>
        /// The sleeping seconds after the mailbox is enabled.
        /// </summary>
        public const string SleepSecondsAfterEnableMailbox = "SleepSecondsAfterEnableMailbox";

        /// <summary>
        /// This value specifies the ID that the client wants associated with the created logon.
        /// </summary>
        public const byte LoginId = 0x00;

        /// <summary>
        /// This index specifies the location in the Server object handle table where the handle for the input Server object is stored.
        /// </summary>
        public const byte InputHandleIndex = 0x00;

        /// <summary>
        /// This index specifies the location in the Server object handle table where the handle for the output Server object will be stored
        /// </summary>
        public const byte OutputHandleIndex = 0x0;
        #endregion Const Global variables
    }
}