namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    ///  An interface serializable object must implement.
    /// </summary>
    public interface IStreamSerializable
    {
        /// <summary>
        /// Serialize object to a FastTransferStream
        /// </summary>
        /// <returns>A FastTransferStream contains the serialized object</returns>
        FastTransferStream Serialize();
    }
}