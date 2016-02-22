namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-ASCMD SUT control adapter interface
    /// </summary>
    public interface IMS_OXWSCORESUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Sets junk email sender to blocked sender list.
        /// </summary>
        /// <param name="itemSenter">The sender of mailbox.</param>
        [MethodHelp("Log on to the server (serverComputerName) with the specified user account (userName, userPassword, userDomain), " +
            "and set junk email sender to blocked sender list. ")]
        string GetMailboxJunkEmailConfiguration();
    }
}