//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_SHDACCWS
{
    using System;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This adapter interface definition of MS-SHDACCWS 
    /// </summary>
    public interface IMS_SHDACCWSAdapter : IAdapter
    {
        #region Interact with versionsService

        /// <summary>
        /// Specifies whether a co-authoring transition request was made for a document.
        /// </summary>
        /// <param name="id">The identifier(Guid) of the document in the server.</param>
        /// <returns>Whether a co-authoring transition request was made for a document.</returns>
        bool IsOnlyClient(Guid id);

        #endregion
    }
}