namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// MS-OXWSUSRCFG SUT control adapter interface.
    /// </summary>
    public interface IMS_OXWSUSRCFGSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Log on to a mailbox with a specified user account and create an user configuration objects on inbox folder.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="password">Password of the user.</param>
        /// <param name="domain">Domain of the user.</param>
        /// <param name="userConfigurationName">Name of the user configuration object.</param>
        /// <returns>If the folder is cleaned up successfully, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and create an user configuration object(userConfigurationName) on inbox folder." +
            " If the operation succeeds, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool CreateUserConfiguration(string userName, string password, string domain, string userConfigurationName);
    }
}
