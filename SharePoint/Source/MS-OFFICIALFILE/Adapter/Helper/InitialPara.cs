namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    /// <summary>
    /// Parameters for InitializeService operation.
    /// </summary>
    public class InitialPara
    {
        /// <summary>
        /// Gets or sets the URL of official file service.
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Gets or sets the property name of user in configuration file.
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// Gets or sets the property name of password in configuration file.
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Gets or sets the property name of domain in configuration file.
        /// </summary>
        public string Domain { get; set; }
    }
}