namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-ASCMD SUT control adapter interface
    /// </summary>
    public interface IMS_OXWSCORESUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Gets junk email sender in blocked sender list.
        /// </summary>
        /// <param name="Organizer">The use account of organizer.</param>

        [MethodHelp("Get junk email sender in blocked sender list of the use account (Organizer). ")]
        string GetMailboxJunkEmailConfiguration(string Organizer);
    }
}