//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
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
