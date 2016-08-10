namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// PropertyValueNode with both property Type and property value
    /// </summary>
    public class AddressBookTypedPropertyValue : AddressBookPropertyValue
    {
        /// <summary>
        /// A 16-bit unsigned integer that specifies the data Type of the property value.
        /// </summary>
        private ushort propertyType;
        
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
            int size = sizeof(byte) * 2;
            size += base.Size();
            return size;
        }

        /// <summary>
        /// Parse bytes in context into TypedPropertyValueNode
        /// </summary>
        /// <param name="context">The value of Context</param>
        public override void Parse(Context context)
        {
            // There's no bytes for PropertyType
            if (context.AvailBytes() < sizeof(ushort))
            {
                return;
            }
            else
            {
                // Parse PropertyType and assign it to context's current PropertyType
                context.CurProperty.Type = (PropertyType)BitConverter.ToUInt16(context.PropertyBytes, context.CurIndex);
                context.CurIndex += 2;
                this.PropertyType = (ushort)context.CurProperty.Type;

                // Parse property value
                base.Parse(context);
            }
        }
    }
}