namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    /// <summary>
    /// This class represents a type that can only set to null value.
    /// </summary>
    public sealed class Null
    {
        /// <summary>
        /// Prevents a default instance of the Null class from being created
        /// </summary>
        private Null()
        {
        }

        /// <summary>
        /// Gets the value null.
        /// </summary>
        /// <value>the only value null.</value>
        public static Null Value
        {
            get
            {
                return null;
            }
        }
    }
}