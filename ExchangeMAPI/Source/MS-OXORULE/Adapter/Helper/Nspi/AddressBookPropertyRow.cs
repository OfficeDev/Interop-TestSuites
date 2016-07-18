namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The AddressBookPropertyRow structure contains a list of property values without including the property tags that correspond to the property values.
    /// </summary>
    public struct AddressBookPropertyRow
    {
        /// <summary>
        /// A byte that indicates whether all property values are present and without error in the ValueArray field.
        /// </summary>
        public byte Flag;

        /// <summary>
        /// An array of variable-sized structures that contains the property values.
        /// </summary>
        public List<AddressBookPropertyValue> ValueArray;

        /// <summary>
        /// Parse the AddressBookPropertyRow structure.
        /// </summary>
        /// <param name="rawBuffer">The raw data returned from server.</param>
        /// <param name="propTagArray">The list of property tags.</param>
        ///  <param name="index">The start index.</param>
        /// <returns>Return an instance of AddressBookPropertyRow.</returns>
        public static AddressBookPropertyRow Parse(byte[] rawBuffer, LargePropTagArray propTagArray, ref int index)
        {
            AddressBookPropertyRow addressBookPropertyRow = new AddressBookPropertyRow();
            addressBookPropertyRow.Flag = rawBuffer[index];
            index++;
            addressBookPropertyRow.ValueArray = new List<AddressBookPropertyValue>();

            Context.Instance.PropertyBytes = rawBuffer;
            Context.Instance.CurIndex = index;
            Context.Instance.CurProperty = new Property(Microsoft.Protocols.TestSuites.Common.PropertyType.PtypUnspecified);

            // If the value of the Flags field is set to 0x00: The array contains either a PropertyValue structure, or a TypedPropertyValue structure.
            // If the value of the Flags field is set to 0x01: The array contains either a FlaggedPropertyValue structure, or a FlaggedPropertyValueWithType structure. 
            if (addressBookPropertyRow.Flag == 0x00)
            {
               foreach (PropertyTag propertyTag in propTagArray.PropertyTags)
               {
                   if (propertyTag.PropertyType == 0x0000)
                   {
                        // If the value of the Flags field is set to 0x00: The array contains a TypedPropertyValue structure, if the type of property is PtyUnspecified.
                       AddressBookTypedPropertyValue typedPropertyValue = new AddressBookTypedPropertyValue();
                       typedPropertyValue.Parse(Context.Instance);
                       addressBookPropertyRow.ValueArray.Add(typedPropertyValue);
                       index = Context.Instance.CurIndex;
                   }
                   else
                   {
                       // If the value of the Flags field is set to 0x00: The array contains a PropertyValue structure, if the type of property is specified.
                       Context.Instance.CurProperty.Type = (Microsoft.Protocols.TestSuites.Common.PropertyType)propertyTag.PropertyType;
                        AddressBookPropertyValue propertyValue = new AddressBookPropertyValue();
                       propertyValue.Parse(Context.Instance);
                       addressBookPropertyRow.ValueArray.Add(propertyValue);
                       index = Context.Instance.CurIndex;
                   }
               }
            }
            else if (addressBookPropertyRow.Flag == 0x01)
            {
                foreach (PropertyTag propertyTag in propTagArray.PropertyTags)
                {
                    if (propertyTag.PropertyType == 0x0000)
                    {
                        // If the value of the Flags field is set to 0x01: The array contains a FlaggedPropertyValueWithType structure, if the type of property is PtyUnspecified.
                        AddressBookFlaggedPropertyValueWithType flaggedPropertyValue = new AddressBookFlaggedPropertyValueWithType();
                        flaggedPropertyValue.Parse(Context.Instance);
                        addressBookPropertyRow.ValueArray.Add(flaggedPropertyValue);
                        index = Context.Instance.CurIndex;
                    }
                    else
                    {
                        // If the value of the Flags field is set to 0x01: The array contains a FlaggedPropertyValue structure, if the type of property is specified.
                        Context.Instance.CurProperty.Type = (Microsoft.Protocols.TestSuites.Common.PropertyType)propertyTag.PropertyType;
                        AddressBookFlaggedPropertyValue propertyValue = new AddressBookFlaggedPropertyValue();
                        propertyValue.Parse(Context.Instance);
                        addressBookPropertyRow.ValueArray.Add(propertyValue);
                        index = Context.Instance.CurIndex;
                    }
                }
            }

            return addressBookPropertyRow;
        }
    }
}