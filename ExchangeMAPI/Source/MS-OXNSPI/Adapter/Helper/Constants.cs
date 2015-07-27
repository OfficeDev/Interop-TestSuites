//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    /// <summary>
    /// Constants used in MS-OXNSPI.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Size of FlatUID_r structure in byte.
        /// </summary>
        public const int FlatUIDByteSize = 16;

        /// <summary>
        /// The maximum number of rows for the NspiGetMatches method to return in a restricted address book container.
        /// </summary>
        public const uint GetMatchesRequestedRowNumber = 5000;

        /// <summary>
        /// The maximum number of rows for the NspiQueryRows method to return in a restricted address book container.
        /// </summary>
        public const uint QueryRowsRequestedRowNumber = 5000;

        /// <summary>
        /// A string which specifies a user name which doesn't exist.
        /// </summary>
        public const string UnresolvedName = "XXXXXX";

        /// <summary>
        /// A CodePage that server does not recognize (0xFFFFFFFF).
        /// </summary>
        public const string UnrecognizedCodePage = "4294967295";

        /// <summary>
        /// A Minimal Entry ID that server does not recognize (0xFFFFFFFF).
        /// </summary>
        public const string UnrecognizedMID = "4294967295";
    }
}