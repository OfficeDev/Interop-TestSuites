//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

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