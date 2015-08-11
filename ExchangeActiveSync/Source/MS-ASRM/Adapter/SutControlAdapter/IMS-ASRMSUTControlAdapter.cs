namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-ASRM SUT control adapter interface.
    /// </summary>
    public interface IMS_ASRMSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Configure the SSL setting in SUT.
        /// </summary>
        /// <param name="serverComputerName">The computer name of the SUT.</param>
        /// <param name="userName">The name of the user used to communicate with server who is in Administrators group of SUT.</param>
        /// <param name="userPassword">The password of the user used to communicate with server.</param>
        /// <param name="userDomain">The domain of the user used to communicate with server.</param>
        /// <param name="enableSSL">If true, SSL setting in SUT should be enabled; otherwise, it should be disabled.</param>
        /// <returns>If succeed, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to the SUT (serverComputerName) with the specified user account (userName, userPassword, userDomain). " +
            "If (enableSSL) is true, enable the SSL setting of ActiveSync; otherwise, disable the SSL setting. " +
            "If the operation succeeds, enter \"true\"; otherwise, enter \"false\".")]
        bool ConfigureSSLSetting(string serverComputerName, string userName, string userPassword, string userDomain, bool enableSSL);
    }
}