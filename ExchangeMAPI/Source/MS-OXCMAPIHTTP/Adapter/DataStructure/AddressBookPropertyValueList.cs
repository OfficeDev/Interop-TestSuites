namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The AddressBookPropertyValueList structure contains a list of properties and their value.
    /// </summary>
    public struct AddressBookPropertyValueList
    {
        /// <summary>
        /// An unsigned integer that specifies the number of structures contained in the PropertyValues field.
        /// </summary>
        public uint PropertyValueCount;

        /// <summary>
        /// An array of AddressBookTaggedPropertyValue structures, each of which specifies a property and its value.
        /// </summary>
        public AddressBookTaggedPropertyValue[] PropertyValues;

        /// <summary>
        /// Parse the AddressBookPropertyValueList structure.
        /// </summary>
        /// <param name="rawData">The raw data of response buffer.</param>
        /// <param name="index">The start index.</param>
        /// <returns>Instance of the AddressBookPropertyValueList.</returns>
        public static AddressBookPropertyValueList Parse(byte[] rawData, ref int index)
        {
            AddressBookPropertyValueList addressBookPropValueList = new AddressBookPropertyValueList();
            addressBookPropValueList.PropertyValueCount = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            Context.Instance.PropertyBytes = rawData;
            Context.Instance.CurIndex = index;
            Context.Instance.CurProperty = new Property(PropertyType.PtypUnspecified);

            addressBookPropValueList.PropertyValues = new AddressBookTaggedPropertyValue[addressBookPropValueList.PropertyValueCount];
            for (int i = 0; i < addressBookPropValueList.PropertyValueCount; i++)
            {
                // Parse the AddressBookTaggedPropertyValue from the response buffer.
                AddressBookTaggedPropertyValue taggedPropertyValue = new AddressBookTaggedPropertyValue();
                taggedPropertyValue.Parse(Context.Instance);
                addressBookPropValueList.PropertyValues[i] = taggedPropertyValue;
            }

            index = Context.Instance.CurIndex;

            return addressBookPropValueList;
        }
    }
}