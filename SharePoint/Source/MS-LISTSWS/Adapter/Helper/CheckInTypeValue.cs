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
    /// The values of various check in types.
    /// </summary>
    public struct CheckInTypeValue
    {
        /// <summary>
        /// A string indicates the value which means Minor check in
        /// </summary>
        public const string MinorCheckIn = "0";

        /// <summary>
        /// A string indicates the value which means Major check in
        /// </summary>
        public const string MajorCheckIn = "1";

        /// <summary>
        /// A string indicates the value which means Overwrite check in
        /// </summary>
        public const string OverwriteCheckIn = "2";

        /// <summary>
        /// A string indicates the value which is invalid for Check in Type value in protocol SUT
        /// </summary>
        public const string InvalidValue = "-1";
    }
}