namespace Microsoft.Protocols.TestSuites.MS_ADMINS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT adapter interface.
    /// </summary>
    public interface IMS_ADMINSSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Get the value of a property in the specified site collection.
        /// </summary>
        /// <param name="url">The url of the specified site collection.</param>
        /// <param name="proName">The name of the property, the possible values are 
        /// following: Title, Description, WebTemplate,  OwnerName, OwnerEmail, PortalUrl, PortalName.</param>
        /// <returns>The value of the property.</returns>
        [MethodHelp("Get the value of a property in the specified site collection. The possible proName values are the following: Title, Description, WebTemplate, OwnerName, OwnerEmail, PortalUrl, and PortalName. Entering a null value will fail the action.")]
        string GetSiteProperty(string url, string proName);

        /// <summary>
        /// If user profile service is implemented by server, this method is used to disable or enable the user profile service in the server.
        /// If user profile service is not implemented by server, this method always returns true.
        /// </summary>
        /// <param name="setDisabled">Input if the user profile service is set to Disabled. True represents setting the user profile service disabled and false represents setting it started.</param>
        /// <returns>Returns if the method is succeed.</returns>
        [MethodHelp("Set the user profile service on the server. The possible setDisabled values are True or False. Setting the value to True will disable the service; and \"False\" will enable the service.")]
        bool SetUserProfileService(bool setDisabled);
    }
}