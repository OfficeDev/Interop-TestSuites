namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    /// <summary>
    /// The request of this protocol.
    /// </summary>
    public class Request
    {
        /// <summary>
        /// The parameter of http request Url.
        /// </summary>
        private string parameter;

        /// <summary>
        /// The body of http request.
        /// </summary>
        private string content;

        /// <summary>
        /// The slug http header.
        /// </summary>
        private string slug;

        /// <summary>
        /// The ETag http header.
        /// </summary>
        private string etag;

        /// <summary>
        /// The http method that used in update request.
        /// </summary>
        private UpdateMethod updateMethod;

        /// <summary>
        /// The Accept http header.
        /// </summary>
        private string accept;

        /// <summary>
        /// The ContentType http header.
        /// </summary>
        private string contentType;

        /// <summary>
        /// Gets or sets the parameter of http request Url.
        /// </summary>
        public string Parameter
        {
            get
            {
                return this.parameter;
            }

            set
            {
                this.parameter = value;
            }
        }

        /// <summary>
        /// Gets or sets the body of http request.
        /// </summary>
        public string Content
        {
            get
            {
                return this.content;
            }

            set
            {
                this.content = value;
            }
        }

        /// <summary>
        /// Gets or sets the slug http header.
        /// </summary>
        public string Slug
        {
            get
            {
                return this.slug;
            }

            set
            {
                this.slug = value;
            }
        }

        /// <summary>
        /// Gets or sets the ETag http header.
        /// </summary>
        public string ETag
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
        /// Gets or sets the http method that used in update request.
        /// </summary>
        public UpdateMethod UpdateMethod
        {
            get
            {
                return this.updateMethod;
            }

            set
            {
                this.updateMethod = value;
            }
        }

        /// <summary>
        /// Gets or sets the Accept http header.
        /// </summary>
        public string Accept
        {
            get
            {
                return this.accept;
            }

            set
            {
                this.accept = value;
            }
        }

        /// <summary>
        /// Gets or sets the ContentType http header.
        /// </summary>
        public string ContentType
        {
            get
            {
                return this.contentType;
            }

            set
            {
                this.contentType = value;
            }
        }
    }
}