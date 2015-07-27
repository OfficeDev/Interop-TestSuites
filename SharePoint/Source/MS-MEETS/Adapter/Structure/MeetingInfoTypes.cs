//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using System;

    /// <summary>
    /// Specifies what information should be returned when calling GetMeetingsInformation 
    /// </summary>
    [Flags]
    public enum MeetingInfoTypes : uint
    {
        /// <summary>
        /// 0x1, queries whether the user has permission to create meeting workspaces on this Web site.
        /// </summary>
        AllowCreate = 1,

        /// <summary>
        /// 0x2, queries for the site template languages supported.
        /// </summary>
        QueryLanguages = 2,

        /// <summary>
        /// 0x4, queries for the list of site templates that are supported.
        /// </summary>
        QueryTemplates = 4,

        /// <summary>
        /// 0x8, queries other status values of a workspace.
        /// </summary>
        QueryOthers = 8
    }
}