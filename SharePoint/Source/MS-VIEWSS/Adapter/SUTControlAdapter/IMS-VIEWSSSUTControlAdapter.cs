//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The MS-VIEWSS SUT Control Adapter interface.
    /// </summary>
    public interface IMS_VIEWSSSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Used to get the count of the list items in the specified view.
        /// </summary>
        /// <param name="listGuid">A specified list GUID in the server.</param>
        /// <param name="viewGuid">A specified view GUID in the server.</param>
        /// <returns>The count of the list items in the specified view.</returns>
        [MethodHelp("Enter the count of the items listed in the specified view (viewGuid) in the specified list (listGuid).")]
        int GetItemsCount(string listGuid, string viewGuid);

        /// <summary>
        /// Used to get the GUID of the specified list.
        /// </summary>
        /// <param name="listDisplayName">A specified list display name in the server.</param>
        /// <returns>The GUID of the specified list.</returns>
        [MethodHelp("Enter the GUID for the specified list (listDisplayName).")]
        string GetListGuidByName(string listDisplayName);

        /// <summary>
        /// Used to retrieve name of the default view for a specified list.
        /// </summary>
        /// <param name="listDisplayName">A specified list display name in the server.</param>
        /// <returns>Name of the default view for the specified list.</returns>
        [MethodHelp("Enter the display name of the default view for the specified list (listDisplayName).")]
        string GetListAndView(string listDisplayName);
    }
}
