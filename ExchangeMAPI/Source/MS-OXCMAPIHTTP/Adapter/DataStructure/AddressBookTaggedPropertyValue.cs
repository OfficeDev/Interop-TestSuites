namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The AddressBookTaggedPropertyValue includes property type, property identifier and property value.
    /// </summary>
    public struct AddressBookTaggedPropertyValue
    {
        /// <summary>
        /// An unsigned integer that identifies the data type of the property value.
        /// </summary>
        public uint? PropertyType;

        /// <summary>
        /// An unsigned integer that identifies the property.
        /// </summary>
        public uint? PropertyId;

        /// <summary>
        /// A property value.
        /// </summary>
        public AddressBookPropertyValue PropertyValue;

        /// <summary>
        /// Parse the AddressBookTaggedPropertyValue structure.
        /// </summary>
        /// <param name="rawBuffer">The raw data returned from server.</param>
        ///  <param name="index">The start index.</param>
        /// <param name="propertyType">The property's type.</param>
        /// <returns>Return an instance of AddressBookTaggedPropertyValue.</returns>
        public static AddressBookTaggedPropertyValue Parse(byte[] rawData, ref int index)
        {
            AddressBookTaggedPropertyValue addressBookTaggedPropertyValue = new AddressBookTaggedPropertyValue();
            Context.Instance.PropertyBytes = rawData;
            Context.Instance.CurIndex = index;

            addressBookTaggedPropertyValue.PropertyType = BitConverter.ToUInt16(rawData, index);
            addressBookTaggedPropertyValue.PropertyId = BitConverter.ToUInt16(rawData, index);

            addressBookTaggedPropertyValue.PropertyValue = AddressBookPropertyValue.Parse(rawData, ref index, (PropertyType)addressBookTaggedPropertyValue.PropertyType);

            index = Context.Instance.CurIndex;

            return addressBookTaggedPropertyValue;
        }
    }
}
