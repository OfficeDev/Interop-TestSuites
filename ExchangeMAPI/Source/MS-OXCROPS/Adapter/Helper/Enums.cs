namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    #region Response types
    /// <summary>
    /// Response types.
    /// </summary>
    public enum RopResponseType
    {
        /// <summary>
        /// Success response.
        /// </summary>
        SuccessResponse,

        /// <summary>
        /// Failure response.
        /// </summary>
        FailureResponse,

        /// <summary>
        /// Response that without any specific request.
        /// </summary>
        Response,

        /// <summary>
        /// Null destination failure response.
        /// </summary>
        NullDestinationFailureResponse,

        /// <summary>
        /// Redirect response.
        /// </summary>
        RedirectResponse,

        /// <summary>
        /// The return code of EcDoRpcExt2 is not zero.
        /// </summary>
        RPCError
    }
    #endregion
}