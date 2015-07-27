//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    /// <summary>
    /// This class contains the constant values
    /// </summary>
    public static class ConstValues
    {
        /// <summary>
        /// An unsigned integer indicates the code page in which text data is sent if Unicode format is not requested by the client on subsequent calls using this Session Context, as specified in MS-OXCRPC section 3.1.4.11 EcDoConnectEx Method.
        /// </summary>
        public const uint CodePageId = 1252;

        /// <summary>
        /// A Default value for Locale. 1033 indicates "en-us".
        /// </summary>
        public const uint DefaultLocale = 0x00000409;

        /// <summary>
        /// An integer indicates a value for reserved field. This value can be any 4-bytes value, as specified in MS-OXCRPC section 3.1.4.11.2.2 EcDoConnectEx method and section 3.1.4.12.2.2 EcDoRpcExt2 method.
        /// </summary>
        public const int ReserveDefault = 0;

        /// <summary>
        /// The flag value of connection
        /// </summary>
        public const uint ConnectionFlag = 0x00000001;

        /// <summary>
        /// The client information used for web HTTP header. 
        /// </summary>
        public const string ClientInfo = "{A7A47AAD-233C-412B-9D10-DDE9108FEBD7}-5";

        /// <summary>
        /// The client application value
        /// </summary>
        public const string ClientApplication = "Outlook/15.00.0856.000";

        /// <summary>
        /// The ROP Logon ID
        /// </summary>
        public const byte LogonId = 0x0;

        /// <summary>
        /// The default output handle used in output ROPS response
        /// </summary>
        public const uint DefaultOutputHandle = 0xFFFFFFFF;

        /// <summary>
        /// The maximum number of rows for the QueryRows method to return in a restricted address book container.
        /// </summary>
        public const uint QueryRowsRequestedRowNumber = 5000;

        /// <summary>
        /// The maximum number of rows for the NspiGetMatches method to return in a restricted address book container.
        /// </summary>
        public const uint GetMatchesRequestedRowNumber = 5000;
    }
}