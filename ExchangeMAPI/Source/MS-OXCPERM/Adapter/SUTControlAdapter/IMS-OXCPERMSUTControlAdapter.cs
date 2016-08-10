namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT Control Adapter interface.
    /// </summary>
    public interface IMS_OXCPERMSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Gets the free/busy status appointment information for User2 (as specified in ptfconfig) through testUserName's account.
        /// </summary>
        /// <param name="testUserName">The user who gets the free/busy status information.</param>
        /// <param name="password">The testUserName's password.</param>
        /// <returns>
        /// <para>"0": means "FreeBusy", which indicates brief information about the appointments on the calendar;</para>
        /// <para>"1": means "Detailed", which indicates detailed information about the appointments on the calendar;</para>
        /// <para>"2": means the appointment free/busy information can't be viewed or error occurs, which indicates the user has no permission to get information about the appointments on the calendar;</para>
        /// </returns>
        [MethodHelp(@"Log in as testUserName to get the free/busy information of User2. The value of User2 is defined in the ""AdminUserName"" property in the MS-OXCPERM_TestSuite.deployment.ptfconfig.\r\n"
            + "Log in as User2 to view the calendar and create an appointment. Note: It is not necessary to create an appointment if User2's calendar already has one.\r\n"
            + "Log in as testUserName to view the appointment on User2's calendar.\r\n"
            + "testUserName: Use this account to open User2's calendar.\r\n"
            + "password: The password for testUserName. \r\n"
            + "The return value is a string type, which indicates the free/busy status or error.\r\n"
            + "\"0\": means \"FreeBusy\", which indicates brief information about the appointments on the calendar;\r\n"
            + "\"1\": means \"Detailed\", which indicates detailed information about the appointments on the calendar;\r\n"
            + "\"2\": means the free/busy information cannot be viewed or an error has occured, which indicates that the user has no permission to get information about the appointments on the calendar.\r\n"
            + "If the server is Microsoft Exchange, the status can be viewed from the Calendar folder in Microsoft Outlook or Microsoft Outlook Web App.")]
        string GetUserFreeBusyStatus(string testUserName, string password);
    }
}