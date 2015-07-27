//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    /// <summary>
    /// The folder type 
    /// </summary>
    public enum FolderTypeEnum
    {
        /// <summary>
        /// Common folder type 
        /// </summary>
        CommonFolderType, 

        /// <summary>
        /// Calendar folder type, special folder type.
        /// </summary>
        CalendarFolderType, 

        /// <summary>
        /// Nonexistent folder type, special folder type.
        /// </summary>
        NonexistentFolderType
    }     
        
    /// <summary>
    /// The permission type
    /// </summary>
    public enum PermissionTypeEnum
    {
        /// <summary>
        /// FreeBusyDetailed flag indicates that the server MUST allow the specified user's client to retrieve detailed information 
        /// about the appointments on the calendar through the Availability Web Service protocol, as specified in [MS-OXWAVLS]. 
        /// </summary>
        FreeBusyDetailed,

        /// <summary>
        /// FreeBusySimple flag indicates that the server MUST allow the specified user's client to retrieve information through the Availability Web Service protocol,
        /// as specified in [MS-OXWAVLS]. 
        /// </summary>
        FreeBusySimple,

        /// <summary>
        /// This flag indicates if the server MUST allow the specified user's client to see the folder in the folder hierarchy table 
        /// and request a handle for the folder by using a RopOpenFolder request, as specified in [MS-OXCFOLD].
        /// </summary>
        FolderVisible,

        /// <summary>
        /// This flag indicates if the server MUST include the specified user in any list of administrative contacts associated with the folder. 
        /// </summary>
        FolderContact,

        /// <summary>
        /// This flag indicates if the server MUST allow the specified user's client to modify properties set on the folder itself, including the folder permissions. 
        /// </summary>
        FolderOwner,

        /// <summary>
        /// This flag indicates if the server MUST allow the specified user's client to create new folders within the folder. 
        /// </summary>
        CreateSubFolder,

        /// <summary>
        /// This flag indicates if the server MUST allow the specified user's client to delete any Message object in the folder. 
        /// </summary>
        DeleteAny,

        /// <summary>
        /// This flag indicates if the server MUST allow the specified user's client to modify any Message object in the folder. 
        /// </summary>
        EditAny,

        /// <summary>
        /// This flag indicates if the server MUST allow the specified user's client to delete any Message object in the folder that was created by that user. 
        /// </summary>
        DeleteOwned,

        /// <summary>
        /// This flag indicates if the server MUST allow the specified user's client to modify any Message object in the folder that was created by that user. 
        /// </summary>
        EditOwned,

        /// <summary>
        /// This flag indicates if the server MUST allow the specified user's client to create new Message objects in the folder. 
        /// </summary>
        Create,

        /// <summary>
        /// This flag indicates if the server MUST allow the specified user's client to read any Message object in the folder. 
        /// </summary>
        ReadAny,

        /// <summary>
        /// This flag indicates PidTagMemberRights reserved bits are set to 1.
        /// </summary>
        Reserved20Permission,
    }
}