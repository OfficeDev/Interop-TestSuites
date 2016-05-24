namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The AddressBookFlaggedPropertyValue includes a flag to indicate whether the value was successfully retrieved or not.
    /// </summary>
    public struct AddressBookFlaggedPropertyValue
    {
        /// <summary>
        /// An unsigned integer that determines what is conveyed in the PropertyValue field.
        /// </summary>
        public uint? Flag;

        /// <summary>
        /// A property value.
        /// </summary>
        public AddressBookPropertyValue PropertyValue;

        /// <summary>
        /// Parse the AddressBookFlaggedPropertyValue structure.
        /// </summary>
        /// <param name="rawBuffer">The raw data returned from server.</param>
        ///  <param name="index">The start index.</param>
        /// <returns>Return an instance of AddressBookFlaggedPropertyValue.</returns>
        public static AddressBookFlaggedPropertyValue Parse(byte[] rawData, ref int index)
        {
            AddressBookFlaggedPropertyValue addressBookFlaggedPropertyValue = new AddressBookFlaggedPropertyValue();
            addressBookFlaggedPropertyValue.Flag = rawData[index++];
            Context.Instance.PropertyBytes = rawData;
            Context.Instance.CurIndex = index;
            Context.Instance.CurProperty = new Property(PropertyType.PtypUnspecified);

            if (addressBookFlaggedPropertyValue.Flag == 0x0)
            {
                addressBookFlaggedPropertyValue.PropertyValue = AddressBookPropertyValue.Parse(rawData, ref index, Context.Instance.CurProperty.Type);
            }

            else if(addressBookFlaggedPropertyValue.Flag == 0xA)
            {
                Context.Instance.CurProperty.Type = PropertyType.PtypErrorCode;
                addressBookFlaggedPropertyValue.PropertyValue = AddressBookPropertyValue.Parse(rawData, ref index, Context.Instance.CurProperty.Type);
            }

            index = Context.Instance.CurIndex;

            return addressBookFlaggedPropertyValue;
        }
    }
}
