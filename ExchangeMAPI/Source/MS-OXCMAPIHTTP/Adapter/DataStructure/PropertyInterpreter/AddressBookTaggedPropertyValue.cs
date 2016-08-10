namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// PropertyValueNode with property tag and property value
    /// </summary>
    [SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1401:FieldsMustBePrivate", Justification = "Disable warning SA1401 because it should not be treated like a property.")]
    public class AddressBookTaggedPropertyValue : AddressBookPropertyValue
    {
        /// <summary>
        /// 16-bit unsigned integer that identifies the data type of the property value.
        /// </summary>
        public ushort PropertyType;

            /// <summary>
        /// A 16-bit unsigned integer that identifies the property.
        /// </summary>
        public ushort PropertyId;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public override byte[] Serialize()
        {
            int index = 0;
            byte[] resultBytes = new byte[this.Size()];
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyType), 0, resultBytes, index, sizeof(ushort));
            index += sizeof(ushort);
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyId), 0, resultBytes, index, sizeof(ushort));
            index += sizeof(ushort);
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
            int size = sizeof(ushort);
            size += sizeof(ushort);
            size += base.Size();
            return size;
        }

        /// <summary>
        /// Parse bytes in context into TaggedPropertyValueNode
        /// </summary>
        /// <param name="context">The value of Context</param>
        public override void Parse(Context context)
        {
            // Parse PropertyType and assign it to context's current PropertyType
            this.PropertyType = (ushort)BitConverter.ToInt16(context.PropertyBytes, context.CurIndex);
            context.CurProperty.Type = (PropertyType)this.PropertyType;
            context.CurIndex += sizeof(ushort);
            this.PropertyId = (ushort)BitConverter.ToInt16(context.PropertyBytes, context.CurIndex);
            context.CurIndex += sizeof(ushort);
            base.Parse(context);           
        }
    }
}