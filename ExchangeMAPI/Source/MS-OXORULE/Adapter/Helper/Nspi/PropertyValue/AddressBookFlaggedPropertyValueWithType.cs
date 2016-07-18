namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// PropertyValueNode with property's value, Type and flag
    /// </summary>
    public class AddressBookFlaggedPropertyValueWithType : AddressBookPropertyValue
    {
        /// <summary>
        /// A 16-bit unsigned integer that specifies the data Type of the property value.
        /// </summary>
        private ushort propertyType;
        
        /// <summary>
        /// An 8-bit unsigned integer. This flag MUST be set one of three possible values: 0x0, 0x1, or 0xA, which determines what is conveyed in the PropertyValue field.
        /// </summary>
        private byte flag;

        /// <summary>
        /// Gets or sets an 8-bit unsigned integer. This flag MUST be set one of three possible values: 0x0, 0x1, or 0xA, which determines what is conveyed in the PropertyValue field.
        /// </summary>
        public byte Flag
        {
            get { return this.flag; }
            set { this.flag = value; }
        }

        /// <summary>
        /// Gets or sets a 16-bit unsigned integer that specifies the data Type of the property value.
        /// </summary>
        public ushort PropertyType
        {
            get { return this.propertyType; }
            set { this.propertyType = value; }
        }

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public override byte[] Serialize()
        {
            int index = 0;
            byte[] resultBytes = new byte[this.Size()];
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyType), 0, resultBytes, index, sizeof(ushort));
            index += 2;
            resultBytes[index++] = this.Flag;
            Array.Copy(base.Serialize(), 0, resultBytes, index, base.Size());
            index += base.Size();
            return resultBytes;
        }

        /// <summary>
        /// Return the Size of this struct.
        /// </summary>
        /// <returns>The Size of this struct.</returns>
        public override int Size()
        {
            int size = sizeof(byte) * 3;
            size += base.Size();
            return size;
        }

        /// <summary>
        /// Parse bytes in context into a FlaggedPropertyValueWithTypeNode
        /// </summary>
        /// <param name="context">The value of Context</param>
        public override void Parse(Context context)
        {
            // Parse Type
            if (context.AvailBytes() < sizeof(ushort))
            {
                throw new ParseException("not well formed FlaggedPropertyValueWithTypeNode with PropertyType missed");
            }
            else
            {
                context.CurProperty.Type = (PropertyType)BitConverter.ToUInt16(context.PropertyBytes, context.CurIndex);
                context.CurIndex += 2;
                this.PropertyType = (ushort)context.CurProperty.Type;
                base.Parse(context);
            }

            // Parse flag
            if (context.AvailBytes() < sizeof(byte))
            {
                throw new ParseException("not well formed FlaggedPropertyValueWithTypeNode with PropertyFlag missed");
            }

            this.Flag = context.PropertyBytes[context.CurIndex++];
            switch (this.Flag)
            {
                // Indicates PropertyValue presents
                case 0:
                    break;

                // Indicates PropertyValue not presents
                case 1:
                    return;

                // Indicates error for this property
                case 0xA:

                    // Define 0xFFFF as a Property that contains error code
                    context.CurProperty.Type = (PropertyType)0xFFFF;
                    break;

                // Not defined value
                default:
                    return;
            }

            // Parse content
            base.Parse(context);
        }
    }
}