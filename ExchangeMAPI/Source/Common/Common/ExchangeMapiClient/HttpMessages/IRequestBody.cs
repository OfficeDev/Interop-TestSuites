namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// A interface of request body for Mailbox Server Endpoint.
    /// </summary>
    public interface IRequestBody
    {
        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>The serialized data to be returned.</returns>
        byte[] Serialize();
    }
}