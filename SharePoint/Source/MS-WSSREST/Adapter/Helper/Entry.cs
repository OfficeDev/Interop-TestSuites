namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// The class specifies the list item.
    /// </summary>
    public class Entry
    {
        /// <summary>
        /// The ETag of response.
        /// </summary>
        private string etag;

        /// <summary>
        /// The title of response.
        /// </summary>
        private string title;

        /// <summary>
        /// The id of response.
        /// </summary>
        private string id;

        /// <summary>
        /// The update time of the response.
        /// </summary>
        private DateTime updated;

        /// <summary>
        /// The content of the list item.
        /// </summary>
        private Dictionary<string, string> properties;

        /// <summary>
        /// Gets or sets the ETag of response.
        /// </summary>
        public string Etag
        {
            get
            {
                return this.etag;
            }

            set
            {
                this.etag = value;
            }
        }

        /// <summary>
        /// Gets or sets the title of response.
        /// </summary>
        public string Title
        {
            get
            {
                return this.title;
            }

            set
            {
                this.title = value;
            }
        }

        /// <summary>
        /// Gets or sets the id of response.
        /// </summary>
        public string ID
        {
            get
            {
                return this.id;
            }

            set
            {
                this.id = value;
            }
        }

        /// <summary>
        /// Gets or sets the update time of the response.
        /// </summary>
        public DateTime Updated
        {
            get
            {
                return this.updated;
            }

            set
            {
                this.updated = value;
            }
        }

        /// <summary>
        /// Gets or sets the content of the list item.
        /// </summary>
        public Dictionary<string, string> Properties
        {
            get
            {
                return this.properties;
            }

            set
            {
                this.properties = value;
            }
        }
    }
}