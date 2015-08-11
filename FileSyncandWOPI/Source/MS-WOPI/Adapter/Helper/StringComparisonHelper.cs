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