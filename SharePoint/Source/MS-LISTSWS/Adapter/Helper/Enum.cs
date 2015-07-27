//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    /// <summary>
    /// [MS-WSSFO2] section 2.2.3.11 for the possible values of the BaseType.
    /// </summary>
    public enum BaseType
    {
        /// <summary>
        /// Generic List Type.
        /// </summary>
        Generic_List = 0,

        /// <summary>
        /// Document Library Type.
        /// </summary>
        Document_Library = 1,

        /// <summary>
        /// Discussion board list Type.
        /// </summary>
        Discussion_Board = 3,

        /// <summary>
        /// Survey list Type.
        /// </summary>
        Survey = 4,

        /// <summary>
        /// Issues list Type.
        /// </summary>
        Issues = 5,
    }
}