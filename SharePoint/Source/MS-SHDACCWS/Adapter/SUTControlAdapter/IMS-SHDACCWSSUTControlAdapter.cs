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
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUTControlAdapter's implementation.
    /// </summary>
    public interface IMS_SHDACCWSSUTControlAdapter : IAdapter
    {
        #region Interact with ListsService

        /// <summary>
        /// Set the Co-authoring status for the specified file under the specified Document LibraryName list. 
        /// The specified file is identified by the property "FileIdOfCoAuthoring".
        /// </summary>
        /// <returns>True if the operation success, otherwise false.</returns>
        [MethodHelp("Set the specified co-authoring status for the specified file which is identified by the property \"FileIdOfCoAuthoring\". Enter \"TRUE\" if the operation succeeds; otherwise, enter \"FALSE\".")]
        bool SUTSetCoAuthoringStatus();

        /// <summary>
        /// Set status of exclusive lock to the specified file which is identified by the property "FileIdOfLock".
        /// </summary>
        /// <returns>True if the operation success, otherwise false.</returns>
        [MethodHelp("Set the specified status of the exclusive lock to the specified file which is identified by the property \"FileIdOfLock\". Enter \"TRUE\" if the operation succeeds; otherwise, enter \"FALSE\".")]
        bool SUTSetExclusiveLock();

        #endregion
    }
}