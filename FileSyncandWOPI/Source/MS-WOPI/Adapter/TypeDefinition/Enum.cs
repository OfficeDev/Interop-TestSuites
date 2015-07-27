//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WOPI
{   
    /// <summary>
    /// An enum represents the operation names which are defined in MS-WOPI.
    /// </summary>
    public enum WOPIOperationName
    {   
        /// <summary>
        /// Represents the GetFile operation.
        /// </summary>
        GetFile,

        /// <summary>
        /// Represents the PutFile operation.
        /// </summary>
        PutFile,

        /// <summary>
        /// Represents the Lock operation.
        /// </summary>
        Lock,

        /// <summary>
        /// Represents the CheckFileInfo operation.
        /// </summary>
        CheckFileInfo,

        /// <summary>
        /// Represents the CheckFolderInfo operation.
        /// </summary>
        CheckFolderInfo,

        /// <summary>
        /// Represents the UnLock operation.
        /// </summary>
        UnLock,

        /// <summary>
        /// Represents the EnumerateChildren operation.
        /// </summary>
        EnumerateChildren,

        /// <summary>
        /// Represents the PutRelativeFile operation.
        /// </summary>
        PutRelativeFile,

        /// <summary>
        /// Represents the RefreshLock operation.
        /// </summary>
        RefreshLock,

        /// <summary>
        /// Represents the UnlockAndRelock operation.
        /// </summary>
        UnlockAndRelock,

        /// <summary>
        /// Represents the ExecuteCellStorageRequest operation.
        /// </summary>
        ExecuteCellStorageRequest,

        /// <summary>
        /// Represents the ExecuteCellStorageRelativeRequest operation.
        /// </summary>
        ExecuteCellStorageRelativeRequest,

        /// <summary>
        /// Represents the DeleteFile operation.
        /// </summary>
        DeleteFile,

        /// <summary>
        /// Represents the GetRestrictedLink operation.
        /// </summary>
        GetRestrictedLink,

        /// <summary>
        /// Represents the RevokeRestrictedLink operation.
        /// </summary>
        RevokeRestrictedLink,

        /// <summary>
        /// Represents the ReadSecureStore operation.
        /// </summary>
        ReadSecureStore
    }

    /// <summary>
    /// An enum represents the WOPI resource URL root level formats which is defined in MS-WOPI.
    /// </summary>
    public enum WOPIRootResourceUrlType
    {  
       /// <summary>
       /// Represents the URL is folder level format.
       /// </summary>
       FolderLevel = 0,

       /// <summary>
       /// Represents the URL is file level format.
       /// </summary>
       FileLevel = 1,
    }

    /// <summary>
    /// An enum represents the WOPI resource URL sub-level formats which are defined in MS-WOPI.
    /// </summary>
    public enum WOPISubResourceUrlType
    {
        /// <summary>
        /// Represents the URL is folder children level format, it is based on the WOPIResourceRootUrlFormat.FolderLevel.
        /// </summary>
        FolderChildrenLevel = 0,

        /// <summary>
        /// Represents the URL is file contents level format, it is based on the WOPIResourceRootUrlFormat.FileLevel.
        /// </summary>
        FileContentsLevel = 1,
    }

    /// <summary>
    /// An enum represents the CellStore operation type.
    /// </summary>
    public enum CellStoreOperationType
    {   
        /// <summary>
        /// Represents the CellStore operation is relative type operation and used to add a file.
        /// </summary>
        RelativeAdd,

        /// <summary>
        /// Represents the CellStore operation is relative type operation and used to update a file.
        /// </summary>
        RealativeModified,

        /// <summary>
        /// Represents the CellStore operation is not a relative type operation and perform normal operation, but without adding a file function.
        /// </summary>
        NormalCellStore
    }
}