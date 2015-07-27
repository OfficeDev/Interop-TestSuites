//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using System;

    /// <summary>
    /// Used to specify the transport type for soap messages.
    /// </summary>
    public enum TransportType
    {
        /// <summary>
        /// Indicate the soap transport over http.
        /// </summary>
        HTTP,

        /// <summary>
        /// Indicate the soap message over https.
        /// </summary>
        HTTPS
    }
 
    /// <summary>
    /// The result status code of a SubmitFile WSDL operation.
    /// </summary>
    public enum SubmitFileResultCode
    {
        /// <summary>
        ///  The operation is successful.
        /// </summary>
        Success,

        /// <summary>
        ///  The operation is successful but further action is needed.
        /// </summary>
        MoreInformation,

        /// <summary>
        ///  The operation failed because of invalid configuration.
        /// </summary>
        InvalidRouterConfiguration,

        /// <summary>
        /// The operation failed because of an invalid argument.
        /// </summary>
        InvalidArgument,

        /// <summary>
        /// The operation failed because of an invalid user.
        /// </summary>
        InvalidUser,

        /// <summary>
        /// The operation failed because of a file not found. 
        /// </summary>
        NotFound,

        /// <summary>
        /// The operation failed because of a rejected file.
        /// </summary>
        FileRejected,

        /// <summary>
        /// The operation failed because of an unknown error.
        /// </summary>
        UnknownError
    }

    /// <summary>
    /// The result of processing a legal hold.
    /// </summary>
    public enum HoldProcessingResult
    {
        /// <summary>
        /// The processing of a legal hold is successful.
        /// </summary>
        Success,

        /// <summary>
        /// The processing of a legal hold failed.
        /// </summary>
        Failure,

        /// <summary>
        /// The file has been stored in the default storage location.
        /// </summary>
        InDropOffZone
    }
}