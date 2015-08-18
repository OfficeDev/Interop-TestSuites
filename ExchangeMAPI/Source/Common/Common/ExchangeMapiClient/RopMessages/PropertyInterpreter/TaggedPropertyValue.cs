namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// PropertyValueNode with property tag and property value
    /// </summary>
    [SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1401:FieldsMustBePrivate", Justification = "Disable warning SA1401 because it should not be treated like a property.")]
    public class TaggedPropertyValue : PropertyValue
    {
        /// <summary>
        /// PropertyTag structure giving the PropertyId and PropertyType for the property
        /// </summary>
        public PropertyTag PropertyTag;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public override byte[] Serialize()
        {
            int index = 0;
            byte[] resultBytes = new byte[this.Size()];
            Array.Copy(this.PropertyTag.Serialize(), 0, resultBytes, index, this.PropertyTag.Size());
            index += this.PropertyTag.Size();
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
            int size = this.PropertyTag.Size();
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
            Microsoft.Protocols.TestSuites.Common.PropertyTag p = new PropertyTag();
            context.CurIndex += p.Deserialize(context.PropertyBytes, context.CurIndex);
            context.CurProperty.Type = (PropertyType)p.PropertyType;

            // this.PropertyTag = new PropertyTag();
            this.PropertyTag = p;
            
            // context.CurIndex += this.PropertyTag.Deserialize(context.PropertyBytes, context.CurIndex);
            // context.CurProperty.Type = (PropertyType)this.PropertyTag.PropertyType;
            base.Parse(context);
        }
    }
}