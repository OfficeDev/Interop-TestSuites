namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXWSCORE SUT control adapter interface
    /// </summary>
    public interface IMS_OXWSCORESUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Gets junk email sender in blocked sender list.
        /// </summary>
        /// <param name="UserName">The user account of organizer.</param>
        [MethodHelp("Get junk email sender in blocked sender list of the user account (UserName). ")]
        string GetMailboxJunkEmailConfiguration(string UserName);
    }
}