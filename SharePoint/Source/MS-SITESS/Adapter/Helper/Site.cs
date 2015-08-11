namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    /// <summary>
    /// Represents information about the site collection.
    /// </summary>
    public class Site
    {
        /// <summary>
        /// The absolute URL of the site collection.
        /// </summary>
        private string urlField;

        /// <summary>
        /// The site collection identifier of the site collection.
        /// </summary>
        private string siteIdField;

        /// <summary>
        /// Specifies whether user code is enabled for the site collection.
        /// </summary>
        private string userCodeEnabledField;

        /// <summary>
        /// Gets or sets the absolute URL of the site collection.
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute]
        public string Url
        {
            get
            {
                return this.urlField;
            }

            set
            {
                this.urlField = value;
            }
        }

        /// <summary>
        /// Gets or sets the site collection identifier of the site collection.
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute]
        public string Id
        {
            get
            {
                return this.siteIdField;
            }

            set
            {
                this.siteIdField = value;
            }
        }

        /// <summary>
        /// Gets or sets the value of the userCodeEnabled field.
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute]
        public string UserCodeEnabled
        {
            get
            {
                return this.userCodeEnabledField;
            }

            set
            {
                this.userCodeEnabledField = value;
            }
        }
    }
}