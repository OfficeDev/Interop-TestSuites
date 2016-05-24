namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The AddressBookTypedPropertyValue includes a property type and property value.
    /// </summary>
    public struct AddressBookTypedPropertyValue
    {
        /// <summary>
        /// An unsigned integer that identifies the data type of the property value.
        /// </summary>
        public uint? PropertyType;

        /// <summary>
        /// A property value.
        /// </summary>
        public AddressBookPropertyValue PropertyValue;

        /// <summary>
        /// Parse the AddressBookTypedPropertyValue structure.
        /// </summary>
        /// <param name="rawBuffer">The raw data returned from server.</param>
        ///  <param name="index">The start index.</param>
        /// <returns>Return an instance of AddressBookTypedPropertyValue.</returns>
        public static AddressBookTypedPropertyValue Parse(byte[] rawData, ref int index)
        {
            AddressBookTypedPropertyValue addressBookTypedPropertyValue = new AddressBookTypedPropertyValue();
            Context.Instance.PropertyBytes = rawData;
            Context.Instance.CurIndex = index;

            addressBookTypedPropertyValue.PropertyType = BitConverter.ToUInt16(rawData, index);

            addressBookTypedPropertyValue.PropertyValue = AddressBookPropertyValue.Parse(rawData, ref index, (PropertyType)addressBookTypedPropertyValue.PropertyType);

            index = Context.Instance.CurIndex;

            return addressBookTypedPropertyValue;
        }
    }
}
