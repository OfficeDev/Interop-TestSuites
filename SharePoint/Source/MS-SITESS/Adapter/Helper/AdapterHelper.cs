//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Xml;
    using System.Xml.Schema;
    using System.Xml.Serialization;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to assist MS-SITESSAdapter.
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary>
        /// An object provides logging, assertions, and SUT adapters for test code onto its execution context.
        /// </summary>
        private static ITestSite testsite;

        /// <summary>
        /// The results of XML Schema validation.
        /// </summary>
        private static Collection<ValidationEventArgs> validationInfo = new Collection<ValidationEventArgs>();

        /// <summary>
        /// Gets the results of XML Schema validation.
        /// </summary>
        public static Collection<ValidationEventArgs> ValidationInfo
        {
            get
            {
                return AdapterHelper.validationInfo;
            }
        }

        /// <summary>
        /// Initialize this helper class with ITestSite.
        /// </summary>
        /// <param name="testSiteInstance">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void Initialize(ITestSite testSiteInstance)
        {
            testsite = testSiteInstance;
        }

        /// <summary>
        /// Deserialize an xml string to a Site object.
        /// </summary>
        /// <param name="getSiteResult">The xml string returned from the GetSite operation.</param>
        /// <returns>The deserialized Site object.</returns>
        public static Site SiteResultDeserialize(string getSiteResult)
        {
            XmlSerializer deserializer = new XmlSerializer(typeof(Site));
            Site siteResult;
            using (StringReader s = new StringReader(getSiteResult))
            {
                siteResult = (Site)deserializer.Deserialize(s);
            }

            return siteResult;
        }

        /// <summary>
        /// Compare two string array's elements.
        /// </summary>
        /// <param name="str1">The first string array to be compared.</param>
        /// <param name="str2">The second string array to be compared.</param>
        /// <returns> If the elements are same, return true, otherwise, return false.</returns>
        public static bool CompareStringArrays(string[] str1, string[] str2)
        {
            bool result = false;
            int count = 0;

            // Sort the string array to be compared since the order of each element in it is not cared.
            Array.Sort(str1);
            Array.Sort(str2);

            if (str1.Length != str2.Length)
            {
                return result;
            }

            for (; count < str1.Length; count++)
            {
                if (str1[count].Equals(str2[count++]))
                {
                    continue;
                }
                else
                {
                    return result;
                }
            }

            result = true;
            return result;
        }

        /// <summary>
        /// Deserialize web properties returned by the GetWebProperties method.
        /// </summary>
        /// <param name="prop">Web properties string value.</param>
        /// <param name="itemSpliter">The char value for item split.</param>
        /// <param name="keySpliter">The char value for key and value split.</param>
        /// <returns>A Dictionary instance contains web properties.</returns>
        public static Dictionary<string, string> DeserializeWebProperties(string prop, char itemSpliter, char keySpliter)
        {
            string propStr = prop.Trim(itemSpliter);
            string[] array = propStr.Split(itemSpliter);
            Dictionary<string, string> result = new Dictionary<string, string>();
            foreach (string str in array)
            {
                string keyStr = str.Trim(keySpliter);
                string[] temp = keyStr.Split(keySpliter);
                if (temp.Length != 2)
                {
                    testsite.Assert.Fail("DeserializeWebProperties error!", null);
                }

                result.Add(temp[0], temp[1]);
            }

            return result;
        }

        /// <summary>
        /// Validate whether an xml message follows the schema.
        /// </summary>
        /// <param name="message">The xml message.</param>
        /// <param name="schema">The schema of the message.</param>
        public static void MessageValidation(string message, string schema)
        {
            validationInfo.Clear();
            XmlReaderSettings schemaSettings = new XmlReaderSettings();
            using (StringReader a = new StringReader(schema))
            {
                schemaSettings.Schemas.Add(null, XmlReader.Create(a));
                schemaSettings.ValidationType = ValidationType.Schema;
                schemaSettings.ConformanceLevel = ConformanceLevel.Document;
                schemaSettings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
                schemaSettings.ValidationEventHandler += new ValidationEventHandler(Schema_ValidationEventHandler);

                using (StringReader m = new StringReader(message))
                {
                    XmlReader xmlReader = XmlReader.Create(m, schemaSettings);
                    try
                    {
                        while (xmlReader.Read())
                        {
                        }
                    }
                    catch (XmlException)
                    {
                        // The exception will be handled by parent caller.
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// The callback method that will handle XML schema validation events.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="args"> A ValidationEventArgs containing the event data.</param>
        private static void Schema_ValidationEventHandler(object sender, ValidationEventArgs args)
        {
            validationInfo.Add(args);
        }
    }
}