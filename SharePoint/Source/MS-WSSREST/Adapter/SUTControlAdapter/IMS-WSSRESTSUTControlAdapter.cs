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
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The MS-WSSREST SUT Control Adapter interface.
    /// </summary>
    public interface IMS_WSSRESTSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Get the document library content type id.
        /// </summary>
        /// <param name="documentListName">The document library name.</param>
        /// <returns>The document library content type id.</returns>
        [MethodHelp("Get the document library content type ID.\r\n")]
        string GetDocumentLibraryContentTypeId(string documentListName);

        /// <summary>
        /// Check whether the type of the specified field equals the expect field type.
        /// </summary>
        /// <param name="fieldName">The specified field name.</param>
        /// <param name="expectFieldType">The expect field type.</param>
        /// <returns>True if the type of the specified field name equals the expect field type, otherwise false.</returns>
        [MethodHelp("Check whether the type of the specified field equals the expect field type.\r\n")]
        bool CheckFieldType(string fieldName, string expectFieldType);
    }
}