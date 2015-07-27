//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// The operation type is used to distinguish the operations in MS-WOPI or MS-FSSHTTP.
    /// </summary>
    public enum OperationType
    {
        /// <summary>
        /// The operations defined in the MS-FSSHTTP.
        /// </summary>
        FSSHTTPCellStorageRequest,

        /// <summary>
        /// The ExecuteCellStorageRequest defined in the MS-WOPI.
        /// </summary>
        WOPICellStorageRequest,

        /// <summary>
        /// The ExecuteCellStorageRelativeRequest defined in the MS-WOPI. 
        /// </summary>
        WOPICellStorageRelativeRequest
    }
}