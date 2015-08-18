namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// Property structure.
    /// </summary>
    public class Property
    {
        /// <summary>
        /// Property name.
        /// </summary>
        private string name;

        /// <summary>
        /// Property type.
        /// </summary>
        private PropertyType type;

        /// <summary>
        /// Initializes a new instance of the <see cref="Property"/> class.
        /// </summary>
        /// <param name="propType">Property type.</param>
        public Property(PropertyType propType)
        {
            this.type = propType;
            this.name = string.Empty;
        }

        /// <summary>
        /// Gets or sets property name.
        /// </summary>
        public string Name
        {
            get
            {
                return this.name;
            }

            set
            {
                this.name = value;
            }
        }

        /// <summary>
        /// Gets or sets property type.
        /// </summary>
        public PropertyType Type
        {
            get
            {
                return this.type;
            }

            set
            {
                this.type = value;
            }
        }
    }
}