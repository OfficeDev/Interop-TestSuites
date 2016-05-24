namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The AddressBookPropertyValue includes a property value.
    /// </summary>
    public struct AddressBookPropertyValue
    {
        /// <summary>
        /// A byte that indicates when the PropertyType ([MS-OXCDATA] section 2.11.1) is known to be either PtypString, PtypString8, PtypBinary or PtypMultiple ([MS-OXCDATA] section 2.11.1).
        /// </summary>
        public byte? HasValue;

        /// <summary>
        /// A property value.
        /// </summary>
        public PropertyValue PropertyValue;

        /// <summary>
        /// Parse the AddressBookPropertyValue structure.
        /// </summary>
        /// <param name="rawBuffer">The raw data returned from server.</param>
        ///  <param name="index">The start index.</param>
        /// <param name="propertyType">The property's type.</param>
        /// <returns>Return an instance of AddressBookPropertyValue.</returns>
        public static AddressBookPropertyValue Parse(byte[] rawData, ref int index, PropertyType propertyType)
        {
            AddressBookPropertyValue addressBookPropertyValue = new AddressBookPropertyValue();
            addressBookPropertyValue.HasValue = rawData[index];
            index++;
            Context.Instance.PropertyBytes = rawData;
            Context.Instance.CurIndex = index;
            Context.Instance.CurProperty.Type = propertyType;

            if (!(propertyType == PropertyType.PtypBinary && propertyType == PropertyType.PtypString && propertyType == PropertyType.PtypString8 && propertyType == PropertyType.PtypMultipleBinary && propertyType == PropertyType.PtypMultipleCurrency && propertyType == PropertyType.PtypMultipleFloating32 
                && propertyType == PropertyType.PtypMultipleFloating64 && propertyType == PropertyType.PtypMultipleFloatingTime && propertyType == PropertyType.PtypMultipleGuid && propertyType == PropertyType.PtypMultipleInteger16 && propertyType == PropertyType.PtypMultipleInteger32 
                && propertyType == PropertyType.PtypMultipleInteger64 && propertyType == PropertyType.PtypMultipleString && propertyType == PropertyType.PtypMultipleString8 && propertyType == PropertyType.PtypMultipleTime && addressBookPropertyValue.HasValue == 0x00))
            {
                PropertyValue value = new PropertyValue();
                value.Parse(Context.Instance);
                addressBookPropertyValue.PropertyValue = value;
            }

            index = Context.Instance.CurIndex;

            return addressBookPropertyValue;
        }
    }
}
