//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    /// <summary>
    /// An enumeration that describes type of item. 
    /// </summary>
    public enum Item
    {
        /// <summary>
        /// Represent a meeting request
        /// </summary>
        MeetingRequest,

        /// <summary>
        /// Represent a meeting response
        /// </summary>
        MeetingResponse,

        /// <summary>
        /// Represent a meeting cancellation
        /// </summary>
        MeetingCancellation,

        /// <summary>
        /// Represent a calendar item
        /// </summary>
        CalendarItem
    }
}
