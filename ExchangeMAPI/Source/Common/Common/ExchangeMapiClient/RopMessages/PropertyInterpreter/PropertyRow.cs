namespace Microsoft.Protocols.TestSuites.Common
{
    using System.Collections.Generic;

    /// <summary>
    /// This class represents a row of properties
    /// </summary>
    public class PropertyRow : Node
    {
        /// <summary>
        /// Property values.
        /// </summary>
        private List<PropertyValue> propertyValues;

        /// <summary>
        /// Flag indicates the Type of PropertyRow
        /// </summary>
        private byte flag;

        /// <summary>
        /// Initializes a new instance of the PropertyRow class.
        /// </summary>
        public PropertyRow()
        {
            // Allocate memory for PropertyValues list
            this.propertyValues = new List<PropertyValue>();
        }

        #region Properties

        /// <summary>
        /// Gets or sets property values.
        /// </summary>
        public List<PropertyValue> PropertyValues
        {
            get { return this.propertyValues; }
            set { this.propertyValues = value; }
        }

        /// <summary>
        /// Gets or sets flag indicates the Type of PropertyRow
        /// </summary>
        public byte Flag
        {
            get { return this.flag; }
            set { this.flag = value; }
        }
        #endregion

        /// <summary>
        /// Parse bytes in context into a PropertyRowNode
        /// </summary>
        /// <param name="context">The value of Context</param>
        public override void Parse(Context context)
        {
            // Clear PropertyValues to store parsing result
            this.propertyValues.Clear();

            // Flag indicates the Type of PropertyRow
            this.flag = context.PropertyBytes[context.CurIndex++];

            foreach (Property prop in context.Properties)
            {
                if (context.IsEnd())
                {
                    throw new ParseException("End prematurely");
                }
                
                PropertyValue valueNode = null;

                if (this.flag == 0)
                {
                    // StandardPropertyRow
                    if (prop.Type == PropertyType.PtypUnspecified)
                    {
                        valueNode = new TypedPropertyValue();
                    }
                    else
                    {
                        valueNode = new PropertyValue();
                    }
                }
                else
                {
                    // FlaggedPropertyRow
                    if (prop.Type == PropertyType.PtypUnspecified)
                    {
                        valueNode = new FlaggedPropertyValueWithType();
                    }
                    else
                    {
                        valueNode = new FlaggedPropertyValue();
                    }
                }

                context.CurProperty = new Property((PropertyType)prop.Type);
                valueNode.Parse(context);
                this.propertyValues.Add(valueNode);
            }
        }
    }
}