//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    /// <summary>
    /// This class is used to generate the token by atomic increase the token by 1. 
    /// </summary>
    public static class SequenceNumberGenerator
    {
        /// <summary>
        /// Specify the max token value.
        /// </summary>
        public const double MAXTOKENVALUE = 4294967295;

        /// <summary>
        /// Specify the max sub request ID.
        /// </summary>
        public const uint MAXSUBREQUSTID = 0xFFFFFFFF;

        /// <summary>
        /// Specify the current token value.
        /// </summary>
        private static long token = 0;

        /// <summary>
        /// Specify the current token value.
        /// </summary>
        private static long subRequestID = 0;

        /// <summary>
        /// Specify the current serialNumber value.
        /// </summary>
        private static long serialNumber = 1;

        /// <summary>
        /// This method is used to get the current token value and atomic adding the token by 1.
        /// </summary>
        /// <returns>Returns the current token value.</returns>
        public static uint GetCurrentToken()
        {
            // Thread safety increase the token by 1.
            long currentToken = System.Threading.Interlocked.Increment(ref token);

            // The token value cannot be exceed the value 4294967295.
            if (currentToken > MAXTOKENVALUE)
            {
                throw new System.InvalidOperationException("The token exceeds the max value 4294967295");
            }

            return (uint)currentToken;
        }

        /// <summary>
        /// This method is used to get the current sub request ID and atomic adding the token by 1.
        /// </summary>
        /// <returns>Returns the current sub request ID.</returns>
        public static uint GetCurrentFSSHTTPBSubRequestID()
        {
            // Thread safety increase the token by 1.
            long currentID = System.Threading.Interlocked.Increment(ref subRequestID);

            // The token value cannot be exceed the value 0xFFFFFFFF.
            if (currentID > MAXSUBREQUSTID)
            {
                throw new System.InvalidOperationException("The token exceeds the max value 4294967295");
            }

            return (uint)currentID;
        }

        /// <summary>
        /// This method is used to get the current serial number and atomic adding the token by 1.
        /// </summary>
        /// <returns>Return the current serial number.</returns>
        public static uint GetCurrentSerialNumber()
        {
            // Thread safety increase the token by 1.
            System.Threading.Interlocked.Increment(ref serialNumber);
            return (uint)serialNumber;
        }
    }
}