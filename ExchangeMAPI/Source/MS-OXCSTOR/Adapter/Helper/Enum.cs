//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{    
    /// <summary>
    /// The type of ROP commands referenced in the Open Specification.
    /// </summary>
    public enum ROPCommandType
    {
        /// <summary>
        /// RopLogon to public folder.
        /// </summary>
        RopLogonPublicFolder,

        /// <summary>
        /// RopLogon to private mailbox.
        /// </summary>
        RopLogonPrivateMailbox,

        /// <summary>
       /// Indicates RopGetReceiveFolder
        /// </summary>
        RopGetReceiveFolder,

        /// <summary>
       /// Indicates RopSetReceiveFolder
        /// </summary>
        RopSetReceiveFolder,

        /// <summary>
       /// Indicates RopGetReceiveFolderTable
        /// </summary>
        RopGetReceiveFolderTable,

        /// <summary>
       /// Indicates RopGetStoreState
        /// </summary>
        RopGetStoreState,

        /// <summary>
       /// Indicates RopGetOwningServers
        /// </summary>
        RopGetOwningServers,

        /// <summary>
       /// Indicates RopPublicFolderIsGhosted
        /// </summary>
        RopPublicFolderIsGhosted,

        /// <summary>
       /// Indicates RopLongTermIdFromId
        /// </summary>
        RopLongTermIdFromId,

        /// <summary>
       /// Indicates RopIdFromLongTermId
        /// </summary>
        RopIdFromLongTermId,

        /// <summary>
       /// Indicates RopGetPerUserLongTermIds
        /// </summary>
        RopGetPerUserLongTermIds,

        /// <summary>
       /// Indicates RopGetPerUserGuid
        /// </summary>
        RopGetPerUserGuid,

        /// <summary>
       /// Indicates RopReadPerUserInformation
        /// </summary>
        RopReadPerUserInformation,

        /// <summary>
       /// Indicates RopWritePerUserInformation
        /// </summary>
        RopWritePerUserInformation,

        /// <summary>
       /// Indicates RopGetPropertiesSpecific
        /// </summary>
        RopGetPropertiesSpecific,

        /// <summary>
       /// Indicates RopSetProperties
        /// </summary>
        RopSetProperties,

        /// <summary>
       /// Indicates RopDeleteProperties
        /// </summary>
        RopDeleteProperties,

        /// <summary>
        /// Others ROP commands.
        /// </summary>
        Others
    }
}