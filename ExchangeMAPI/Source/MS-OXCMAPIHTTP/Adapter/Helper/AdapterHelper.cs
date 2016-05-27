namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to assist MS-OXCMAPIHTTPAdapter.
    /// </summary>
    public class AdapterHelper
    {
        #region Variables

        /// <summary>
        /// A GUID string of client instance.
        /// </summary>
        private static string clientInstance;

        /// <summary>
        /// The counter used for request ID.
        /// </summary>
        private static int counter;

        /// <summary>
        /// The current session context cookies for the request.
        /// </summary>
        private static CookieCollection sessionContextCookies;

        /// <summary>
        /// The site which is used to print log information.
        /// </summary>
        private static ITestSite site;

        /// <summary>
        /// Gets or sets the instance of site.
        /// </summary>
        public static ITestSite Site
        {
            get { return AdapterHelper.site; }
            set { AdapterHelper.site = value; }
        }

        /// <summary>
        /// Gets or sets the client instance.
        /// </summary>
        public static string ClientInstance
        {
            get { return clientInstance; }
            set { clientInstance = value; }
        }

        /// <summary>
        /// Gets or sets the counter.
        /// </summary>
        public static int Counter
        {
            get { return counter; }
            set { counter = value; }
        }

        /// <summary>
        /// Gets or sets the current session context cookies for the request.
        /// </summary>
        public static CookieCollection SessionContextCookies
        {
            get
            {
                if (sessionContextCookies == null)
                {
                    sessionContextCookies = new CookieCollection();
                }

                return sessionContextCookies;
            }

            set
            {
                sessionContextCookies = value;
            }
        }

        #endregion Variables

        #region Adapter Help Methods
        /// <summary>
        /// Initialize HTTP Header
        /// </summary>
        /// <param name="requestType">The request type</param>
        /// <param name="clientInstance">The string of the client instance</param>
        /// <param name="counter">The counter</param>
        /// <returns>The web header collection</returns>
        public static WebHeaderCollection InitializeHTTPHeader(RequestType requestType, string clientInstance, int counter)
        {
            WebHeaderCollection webHeaderCollection = new WebHeaderCollection();
            webHeaderCollection.Add("X-ClientInfo", ConstValues.ClientInfo);
            webHeaderCollection.Add("X-RequestId", clientInstance + ":" + counter.ToString());
            webHeaderCollection.Add("X-ClientApplication", ConstValues.ClientApplication);
            webHeaderCollection.Add("X-RequestType", requestType.ToString());
            return webHeaderCollection;
        }

        /// <summary>
        /// Compare whether two AddressBookPropValueLists are equal.
        /// </summary>
        /// <param name="addressBookPropValueList1">The first AddressBookPropertyValueList used to compare.</param>
        /// <param name="addressBookPropValueList2">The second AddressBookPropertyValueList used to compare.</param>
        /// <returns>Returns true if they are equal; otherwise false.</returns>
        public static bool AreTwoAddressBookPropValueListEqual(AddressBookPropertyValueList addressBookPropValueList1, AddressBookPropertyValueList addressBookPropValueList2)
        {
            if (addressBookPropValueList1.PropertyValueCount != addressBookPropValueList2.PropertyValueCount)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length the two AddressBookPropertyValueList are not equal. The length of addressBookPropValueList1 is {0}, the length of addressBookPropValueList2 is {1}.",
                    addressBookPropValueList1.PropertyValueCount,
                    addressBookPropValueList2.PropertyValueCount);

                return false;
            }
            else
            {
                AddressBookTaggedPropertyValue[] propertyValues1 = addressBookPropValueList1.PropertyValues;
                AddressBookTaggedPropertyValue[] propertyValues2 = addressBookPropValueList2.PropertyValues;
                
                for (int i = 0; i < propertyValues1.Length; i++)
                { 
                    if (propertyValues1[i].PropertyId != propertyValues2[i].PropertyId)
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The property ID of property {0} in the two property tag array are not equal. The property ID of propertyTag1 is {1}, the property ID of propertyTag2 is {2}",
                            i,
                            propertyValues1[i].PropertyId,
                            propertyValues2[i].PropertyId);

                        return false;
                    }
                    else if (propertyValues1[i].PropertyType != propertyValues2[i].PropertyType)
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The property type of property {0} in the two property tag array are not equal. The property type of propertyTag1 is {1}, the property type of propertyTag2 is {2}",
                            i,
                            propertyValues1[i].PropertyType,
                            propertyValues2[i].PropertyType);

                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Parse PermanentEntryID structure from byte array.
        /// </summary>
        /// <param name="bytes">The byte array to be parsed.</param>
        /// <returns>An PermanentEntryID structure instance.</returns>
        public static PermanentEntryID ParsePermanentEntryIDFromBytes(byte[] bytes)
        {
            int index = 0;

            // The count size of the buffer
            uint size = BitConverter.ToUInt32(bytes, index);
            index += 4;

            PermanentEntryID entryID = new PermanentEntryID();
            entryID.IDType = bytes[index++];
            entryID.R1 = bytes[index++];
            entryID.R2 = bytes[index++];
            entryID.R3 = bytes[index++];
            byte[] bytesGUID = new byte[16];
            Array.Copy(bytes, index, bytesGUID, 0, 16);
            entryID.ProviderUID = new Guid(bytesGUID);
            index += 16;

            // R4: 4 bytes
            entryID.R4 = (uint)BitConverter.ToInt32(bytes, index);
            index += 4;

            // DisplayType: 4 bytes
            entryID.DisplayTypeString = (DisplayTypeValues)BitConverter.ToInt32(bytes, index);
            index += 4;

            // DistinguishedName: variable 
            entryID.DistinguishedName = System.Text.Encoding.Default.GetString(bytes, index, bytes.Length - index - 1);
            return entryID;
        }

        /// <summary>
        /// Compare whether two LargePropertyTagArrays are equal.
        /// </summary>
        /// <param name="propertyTagArray1">The first LargePropertyTagArray used to compare.</param>
        /// <param name="propertyTagArray2">The second LargePropertyTagArray used to compare.</param>
        /// <returns>Returns true if they are equal; otherwise false.</returns>
        public static bool AreTwoLargePropertyTagArrayEqual(LargePropertyTagArray propertyTagArray1, LargePropertyTagArray propertyTagArray2)
        {
            if (propertyTagArray1.PropertyTagCount != propertyTagArray2.PropertyTagCount)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length the two property tag array are not equal. The length of propertyTagArray1 is {0}, the length of propertyTagArray2 is {1}.",
                    propertyTagArray1.PropertyTagCount,
                    propertyTagArray2.PropertyTagCount);

                return false;
            }
            else
            {
                PropertyTag[] propertyTags1 = propertyTagArray1.PropertyTags;
                PropertyTag[] propertyTags2 = propertyTagArray2.PropertyTags;

                for (int i = 0; i < propertyTags1.Length; i++)
                {
                    if (propertyTags1[i].PropertyId != propertyTags2[i].PropertyId)
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The property ID of property {0} in the two property tag array are not equal. The property ID of propertyTags1 is {1}, the property ID of propertyTags2 is {2}",
                            i,
                            propertyTags1[i].PropertyId,
                            propertyTags2[i].PropertyId);

                        return false;
                    }
                    else if (propertyTags1[i].PropertyType != propertyTags2[i].PropertyType)
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The property type of property {0} in the two property tag array are not equal. The property type of propertyTags1 is {1}, the property type of propertyTags2 is {2}",
                            i,
                            propertyTags1[i].PropertyType,
                            propertyTags2[i].PropertyType);

                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether two AddressBookPropertyRow arrays are equal.
        /// </summary>
        /// <param name="propertyRows1">The first AddressBookPropertyRow array used to compare.</param>
        /// <param name="propertyRows2">The second AddressBookPropertyRow array used to compare.</param>
        /// <returns>Returns true if they are equal; otherwise false.</returns>
        public static bool AreTwoAddressBookPropertyRowEqual(AddressBookPropertyRow[] propertyRows1, AddressBookPropertyRow[] propertyRows2)
        {
            if (propertyRows1.Length != propertyRows2.Length)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of propertyRows1 is {0}, the length of propertyRows2 is {1}.",
                    propertyRows1.Length,
                    propertyRows2.Length);

                return false;
            }
            else
            {
                for (int i = 0; i < propertyRows1.Length; i++)
                {
                    AddressBookPropertyRow propertyRow1 = propertyRows1[i];
                    AddressBookPropertyRow propertyRow2 = propertyRows2[i];

                    if (propertyRow1.Flag != propertyRow2.Flag)
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The value of Flag field of propertyRow1 is {0}, the value of Flag field of propertyRow2 is {1}.",
                            propertyRow1.Flag,
                            propertyRow2.Flag);

                        return false;
                    }
                    else
                    {
                        List<AddressBookPropertyValue> valueArray1 = new List<AddressBookPropertyValue>();
                        valueArray1.AddRange(propertyRow1.ValueArray);
                        List<AddressBookPropertyValue> valueArray2 = new List<AddressBookPropertyValue>(); 
                        valueArray2.AddRange(propertyRow2.ValueArray);

                        if (valueArray1.Count != valueArray2.Count)
                        {
                            site.Log.Add(
                                LogEntryKind.Debug,
                                "The length of valueArray1 is {0}, the length of valueArray2 is {1}.",
                                valueArray1.Count,
                                valueArray2.Count);

                            return false;
                        }
                        else
                        {
                            for (int j = 0; j < valueArray1.Count; j++)
                            {
                                if (valueArray1[j].Value.Length != valueArray2[j].Value.Length)
                                {
                                    site.Log.Add(
                                        LogEntryKind.Debug,
                                        "The length of the first property value is {0}, the length of the second property value is {1}.",
                                        valueArray1[j].Value.Length,
                                        valueArray2[j].Value.Length);

                                    return false;
                                }
                                else
                                {
                                    byte[] valueOfProperty1 = valueArray1[j].Value;
                                    byte[] valueOfProperty2 = valueArray2[j].Value;

                                    for (int k = 0; k < valueOfProperty1.Length; k++)
                                    {
                                        if (valueOfProperty1[k] != valueOfProperty2[k])
                                        {
                                            site.Log.Add(
                                               LogEntryKind.Debug,
                                               "The {0} bit of the first property value is {1}, The {0} bit of the second property value is {2}.",
                                               valueOfProperty1[k],
                                               valueOfProperty2[k]);

                                            return false;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Get the final value of X-ResponseCode header.
        /// </summary>
        /// <param name="responseCode">The value of X-ResponseCode header returned from server.</param>
        /// <returns>Return the final response code.</returns>
        public static uint GetFinalResponseCode(string responseCode)
        {
            if (string.IsNullOrEmpty(responseCode))
            {
                return 0;
            }

            string[] responseCodeValues = responseCode.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            uint finalResponseCode = uint.Parse(responseCodeValues[responseCodeValues.Length - 1]);

            return finalResponseCode;
        }
        #endregion Methods
    }
}