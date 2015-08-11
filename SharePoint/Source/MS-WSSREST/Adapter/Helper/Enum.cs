namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    /// <summary>
    /// The method of http request.
    /// </summary>
    public enum HttpMethod
    {
        /// <summary>
        /// Used in retrieve request.
        /// </summary>
        GET,

        /// <summary>
        /// Used in update request.
        /// </summary>
        PUT,

        /// <summary>
        /// Used in insert request.
        /// </summary>
        POST,

        /// <summary>
        /// Used in delete request.
        /// </summary>
        DELETE,

        /// <summary>
        /// Used in update request.
        /// </summary>
        MERGE
    }

    /// <summary>
    /// The http method of update request.
    /// </summary>
    public enum UpdateMethod
    {
        /// <summary>
        /// Replace the content in the request.
        /// </summary>
        PUT,

        /// <summary>
        /// Merge the content in the request.
        /// </summary>
        MERGE
    }

    /// <summary>
    /// The operation type supported by batch request.
    /// </summary>
    public enum OperationType
    {
        /// <summary>
        /// The insert operation.
        /// </summary>
        Insert,

        /// <summary>
        /// The update operation.
        /// </summary>
        Update,

        /// <summary>
        /// The delete operation.
        /// </summary>
        Delete
    }
}