namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// PropertyValueNode with property's value and error code
    /// </summary>
    public class AddressBookFlaggedPropertyValue : AddressBookPropertyValue
    {
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
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public override byte[] Serialize()
        {
            int index = 0;
            byte[] resultBytes = new byte[this.Size()];
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
            int size = sizeof(byte);
            size += base.Size();
            return size;
        }

        /// <summary>
        /// Parse bytes in context into a FlaggedPropertyValueNode
        /// </summary>
        /// <param name="context">The value of Context</param>
        public override void Parse(Context context)
        {
            this.Flag = context.PropertyBytes[context.CurIndex++];
            switch (this.Flag)
            {
                // PropertyValue presents
                case 0:
                    break;

                // PropertyValue not presents
                case 1:
                    return;

                // Property error code
                case 0xA:

                    // If the Flag is 0x0A, the property type should be PtypErrorCode.
                    context.CurProperty.Type = PropertyType.PtypErrorCode;
                    break;

                // Not defined value
                default:
                    break;
            }
            
            base.Parse(context);
        }
    }
}