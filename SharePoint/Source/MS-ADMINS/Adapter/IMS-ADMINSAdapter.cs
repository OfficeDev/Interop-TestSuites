namespace Microsoft.Protocols.TestSuites.MS_ADMINS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-ADMINS adapter interface.
    /// </summary>
    public interface IMS_ADMINSAdapter : IAdapter
    {
        /// <summary>
        /// Gets or sets the entry point Url of web service operation.
        /// </summary>
        /// <value>The service Url</value>
        string Url
        {
            get;
            set;
        }

        /// <summary>
        /// Creates a Site collection.
        /// </summary>
        /// <param name="url">The absolute URL of the site collection.</param>
        /// <param name="title">The display name of the site collection.</param>
        /// <param name="description">A description of the site collection.</param>
        /// <param name="lcid">The language that is used in the site collection.</param>
        /// <param name="webTemplate">The name of the site template which is used when creating the site collection.</param>
        /// <param name="ownerLogin">The user name of the site collection owner.</param>
        /// <param name="ownerName">The display name of the owner.</param>
        /// <param name="ownerEmail">The e-mail address of the owner.</param>
        /// <param name="portalUrl">The URL of the portal site for the site collection.</param>
        /// <param name="portalName">The name of the portal site for the site collection.</param>
        /// <returns>The CreateSite result.</returns>
        string CreateSite(string url, string title, string description, int? lcid, string webTemplate, string ownerLogin, string ownerName, string ownerEmail, string portalUrl, string portalName);

        /// <summary>
        /// Deletes the specified Site collection.
        /// </summary>
        /// <param name="url">The absolute URL of the site collection which is to be deleted.</param>
        void DeleteSite(string url);

        /// <summary>
        /// Returns information about the languages which are used in the protocol server deployment.
        /// </summary>
        /// <returns>The GetLanguages result.</returns>
        GetLanguagesResponseGetLanguagesResult GetLanguages();
    }
}