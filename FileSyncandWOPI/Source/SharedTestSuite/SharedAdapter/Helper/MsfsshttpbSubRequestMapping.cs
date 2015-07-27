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
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class is used to maintain the sub request token for ms-fsshttpb.
    /// </summary>
    public static class MsfsshttpbSubRequestMapping
    {
        /// <summary>
        /// Dictionary for storing the mapping between the sub request token and sub request type.  
        /// </summary>
        private static Dictionary<int, RequestTypes> msfsshttpbSubRequestIDTypeMapping = new Dictionary<int, RequestTypes>();

        /// <summary>
        /// This method is used to add a mapping between the sub request token and sub request type.
        /// </summary>
        /// <param name="subRequestID">Specify the sub request token.</param>
        /// <param name="subRequestType">Specify the sub request type.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void Add(int subRequestID, Type subRequestType, ITestSite site)
        {
            if (msfsshttpbSubRequestIDTypeMapping.ContainsKey(subRequestID))
            {
                site.Assert.Fail("Duplicate sub-request id in the sub request id and type mapping, possible reason: 1) Does not send unique ms-fsshttpb sub-request id in one ms-fsshttpb request. 2) Does not clean the sub-request mapping. ");
            }

            if (!SubResponseData.SubResponseDataTypeReverseMapping.ContainsKey(subRequestType))
            {
                site.Assert.Fail("Unsupported sub request type " + subRequestType.Name);
            }

            msfsshttpbSubRequestIDTypeMapping.Add(subRequestID, SubResponseData.SubResponseDataTypeReverseMapping[subRequestType]);
        }

        /// <summary>
        /// This method is used to get the sub request type using the specified sub request token.
        /// </summary>
        /// <param name="subRequestID">Specify the sub request token.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        /// <returns>Return the sub request type.</returns>
        public static RequestTypes GetSubRequestType(int subRequestID, ITestSite site)
        {
            if (!msfsshttpbSubRequestIDTypeMapping.ContainsKey(subRequestID))
            {
                site.Assert.Fail("Not record the sub-request id, possible reason: Not using CellRequest::AddSubRequest method to add sub Request.");
            }

            return msfsshttpbSubRequestIDTypeMapping[subRequestID];
        }

        /// <summary>
        /// The method is used to clear all the mappings.
        /// </summary>
        public static void Clear()
        {
            msfsshttpbSubRequestIDTypeMapping.Clear();
        }
    }
}