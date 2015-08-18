namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT adapter interface which is used by test cases in the test suite to set or clear Out of Office state by calling Exchange OOF Web Service. 
    /// </summary>
    public interface IMS_OXWOOFSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Set or clear user's mailbox "Out Of Office" state.
        /// </summary>
        /// <param name="mailAddress">User's email address.</param>
        /// <param name="password">Password of user mailbox.</param>
        /// <param name="isOOF">If true, set the OOF state to be enabled, otherwise set the OOF state to be disabled.</param>
        /// <returns>If the operation succeed then return true, otherwise return false.</returns>
        [MethodHelp("Set Out of Office state for user (mailAddress, password). If the value of parameter isOOF is true, set the OOF state as enabled, otherwise set the OOF state as disabled." +
            "When the OOF state is enabled, the body of the reply message can be blank." +
            "Enter \"true\" if the OOF state is set successfully, otherwise enter \"false\".")]
        bool SetUserOOFSettings(string mailAddress, string password, bool isOOF);
    }
}