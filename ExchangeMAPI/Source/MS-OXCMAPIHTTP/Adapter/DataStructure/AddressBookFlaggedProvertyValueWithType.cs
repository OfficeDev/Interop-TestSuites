namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The AddressBookFlaggedPropertyValueWithType includes both the property type and a flag giving more information about the property value.
    /// </summary>
    public struct AddressBookFlaggedPropertyValueWithType
    {
        /// <summary>
        /// An unsigned integer that identifies the data type of the property value.
        /// </summary>
        public uint? PropertyType;

        /// <summary>
        /// An unsigned integer that determines what is conveyed in the PropertyValue field.
        /// </summary>
        public uint? Flag;

        /// <summary>
        /// A property value.
        /// </summary>
        public AddressBookPropertyValue PropertyValue;

        /// <summary>
        /// Parse the AddressBookFlaggedPropertyValueWithType structure.
        /// </summary>
        /// <param name="rawBuffer">The raw data returned from server.</param>
        ///  <param name="index">The start index.</param>
        /// <returns>Return an instance of AddressBookFlaggedPropertyValueWithType.</returns>
        public static AddressBookFlaggedPropertyValueWithType Parse(byte[] rawData, ref int index)
        {
            AddressBookFlaggedPropertyValueWithType addressBookFlaggedPropertyValueWithType = new AddressBookFlaggedPropertyValueWithType();
            addressBookFlaggedPropertyValueWithType.Flag = rawData[index++];
            Context.Instance.CurProperty.Type = (PropertyType)BitConverter.ToUInt16(Context.Instance.PropertyBytes, Context.Instance.CurIndex);
            Context.Instance.CurIndex += 2;
            addressBookFlaggedPropertyValueWithType.PropertyType = (ushort)Context.Instance.CurProperty.Type;

            if (addressBookFlaggedPropertyValueWithType.Flag == 0x0)
            {
                addressBookFlaggedPropertyValueWithType.PropertyValue = AddressBookPropertyValue.Parse(rawData, ref index, (PropertyType)addressBookFlaggedPropertyValueWithType.PropertyType);
            }

            else if (addressBookFlaggedPropertyValueWithType.Flag == 0xA)
            {
                Context.Instance.CurProperty.Type = (PropertyType)0xFFFF;
                addressBookFlaggedPropertyValueWithType.PropertyValue = AddressBookPropertyValue.Parse(rawData, ref index, Context.Instance.CurProperty.Type);
            }

            index = Context.Instance.CurIndex;

            return addressBookFlaggedPropertyValueWithType;
        }
    }
}
