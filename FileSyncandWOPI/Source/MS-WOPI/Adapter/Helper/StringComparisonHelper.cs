//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class is used to support string comparison feature.
    /// </summary>
    public static class StringComparisonHelper
    {
        /// <summary>
        /// A method used to compare the actual string value whether equal to the expected value by ignoring the case.
        /// </summary>
        /// <param name="expectedValue">A parameter represents the expected string value.</param>
        /// <param name="actualValue">A parameter represents the actual string value.</param>
        /// <param name="testSiteInstance">A parameter represents the ITestSite instance.</param>
        /// <returns>Return 'true' indicating two string values are equal, otherwise those two string values are not equal.</returns>
        public static bool CompareStringValueIgnoreCase(this string expectedValue, string actualValue, ITestSite testSiteInstance)
        {
            if (null == testSiteInstance)
            {
                throw new ArgumentNullException("testSiteInstance");
            }

            bool compareResult = string.Equals(expectedValue, actualValue, StringComparison.OrdinalIgnoreCase);
            testSiteInstance.Log.Add(
                LogEntryKind.Debug,
                "Comparing string values:\r\nActualValue:[{0}]\r\nExpectedValue:[{1}]\r\nResult:[{2}]",
                string.IsNullOrEmpty(actualValue) ? "Null/Empty" : actualValue,
                string.IsNullOrEmpty(expectedValue) ? "Null/Empty" : expectedValue,
                compareResult ? "AreEqual" : "AreNotEqual");

            return compareResult;
        }
    }
}