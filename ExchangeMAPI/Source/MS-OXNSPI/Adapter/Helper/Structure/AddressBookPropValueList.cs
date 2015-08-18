namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The AddressBookPropValueList structure contains a list of properties and their value.
    /// </summary>
    public struct AddressBookPropValueList
    {
        /// <summary>
        /// An unsigned integer that specifies the number of structures contained in the PropertyValues field.
        /// </summary>
        public uint PropertyValueCount;

        /// <summary>
        /// An array of TaggedPropertyValue structures, each of which specifies a property and its value.
        /// </summary>
        public TaggedPropertyValue[] PropertyValues;

        /// <summary>
        /// Parse the AddressBookPropValueList structure.
        /// </summary>
        /// <param name="rawData">The raw data of response buffer.</param>
        /// <param name="index">The start index.</param>
        /// <returns>Instance of the AddressBookPropValueList.</returns>
        public static AddressBookPropValueList Parse(byte[] rawData, ref int index)
        {
            AddressBookPropValueList addressBookPropValueList = new AddressBookPropValueList();
            addressBookPropValueList.PropertyValueCount = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            Context.Instance.PropertyBytes = rawData;
            Context.Instance.CurIndex = index;
            Context.Instance.CurProperty = new Property(PropertyType.PtypUnspecified);

            addressBookPropValueList.PropertyValues = new TaggedPropertyValue[addressBookPropValueList.PropertyValueCount];
            for (int i = 0; i < addressBookPropValueList.PropertyValueCount; i++)
            {
                // Parse the TaggedPropertyValue from the response buffer.
                TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue();
                taggedPropertyValue.Parse(Context.Instance);
                addressBookPropValueList.PropertyValues[i] = taggedPropertyValue;
            }

            index = Context.Instance.CurIndex;

            return addressBookPropValueList;
        }
    }
}