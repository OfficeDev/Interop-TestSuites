namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Defines methods used by test suite to control the SUT
    /// </summary>
    public interface IMS_MEETSSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Make sure there are no meeting workspaces under the specified site.
        /// </summary>
        /// <param name="siteUrl">The site Url</param>
        /// <returns>Returns if the method is succeed.</returns>
        [MethodHelp("Make sure there are no meeting workspaces under the specified site. Entering True in return field indicates that there are no meeting workspaces under the specified site. Entering False indicates that there are some meeting workspaces under the specified site.")]
        bool PrepareTestEnvironment(string siteUrl);
    }
}