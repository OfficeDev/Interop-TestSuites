namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the property that contains long property id.
    /// </summary>
    public class PropertyNameObject
    {
        /// <summary>
        /// A PropertyName object of the property.
        /// </summary>
        private PropertyName propertyName = new PropertyName();

        /// <summary>
        /// The data type of the property.
        /// </summary>
        private PropertyType propertyType;

        /// <summary>
        /// A string value indicates the name of property.
        /// </summary>
        private PropertyNames displayName;

        /// <summary>
        /// Initializes a new instance of the PropertyNameObject class.
        /// </summary>
        /// <param name="displayName">A string value indicates the name of property.</param>
        /// <param name="longId">A unsigned integer value indicates property long ID (LID) of specified property.</param>
        /// <param name="propertySet">A string indicates property set of specified property.</param>
        /// <param name="dataType">The date type of specified property.</param>
        public PropertyNameObject(PropertyNames displayName, uint longId, string propertySet, PropertyType dataType)
        {
            this.displayName = displayName;
            this.propertyName.Kind = 0x00;
            this.propertyName.Guid = new Guid(propertySet).ToByteArray();
            this.propertyName.LID = longId;
            this.propertyType = dataType;
        }

        /// <summary>
        /// Initializes a new instance of the PropertyNameObject class.
        /// </summary>
        /// <param name="displayName">A string indicates display name of specified property.</param>
        /// <param name="name">A string indicates property name of specified property.</param>
        /// <param name="propertySet">A string indicates property set of specified property.</param>
        /// <param name="dataType">The date type of specified property.</param>
        public PropertyNameObject(PropertyNames displayName, string name, string propertySet, PropertyType dataType)
        {
            this.displayName = displayName;
            this.propertyName.Kind = 0x01;
            this.propertyName.Guid = new Guid(propertySet).ToByteArray();
            byte[] nameArray = Common.GetBytesFromUnicodeString(name);
            this.propertyName.Name = nameArray;
            this.propertyName.NameSize = (byte)nameArray.Length;
            this.propertyType = dataType;
        }

        /// <summary>
        /// Gets the PropertyName object of the property.
        /// </summary>
        public PropertyName PropertyName
        {
            get { return this.propertyName; }
        }

        /// <summary>
        /// Gets the data type of the property.
        /// </summary>
        public PropertyType PropertyType
        {
            get { return this.propertyType; }
        }

        /// <summary>
        /// Gets the display name of the property.
        /// </summary>
        public PropertyNames DisplayName
        {
            get { return this.displayName; }
        }
    }
}