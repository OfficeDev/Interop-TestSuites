//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Property Object encapsulated in its ID, Name, Type, and Value. 
    /// </summary>
    public class PropertyObj
    {
        /// <summary>
        /// Property id
        /// </summary>
        private uint propertyID;

        /// <summary>
        /// Property name
        /// </summary>
        private PropertyNames propertyName;

        /// <summary>
        /// Property type
        /// </summary>
        private PropertyType valueType;

        /// <summary>
        /// Object value
        /// </summary>
        private object value = null;

        /// <summary>
        /// Initializes a new instance of the PropertyObj class.
        /// </summary>
        public PropertyObj()
        {
        }

        /// <summary>
        /// Initializes a new instance of the PropertyObj class.
        /// </summary>
        /// <param name="propertyName">NmId of property</param>
        public PropertyObj(PropertyNames propertyName)
        {
            if (PropertyHelper.PropertyTagDic.ContainsKey(propertyName))
            {
                this.PropertyName = propertyName;
                this.propertyID = PropertyHelper.PropertyTagDic[propertyName].PropertyId;
                this.ValueType = (PropertyType)PropertyHelper.PropertyTagDic[propertyName].PropertyType;
            }
        }

        /// <summary>
        /// Initializes a new instance of the PropertyObj class.
        /// </summary>
        /// <param name="propertyName">A property name value</param>
        /// <param name="value">The value of property</param>
        public PropertyObj(PropertyNames propertyName, object value)
        {
            if (PropertyHelper.PropertyTagDic.ContainsKey(propertyName))
            {
                this.PropertyName = propertyName;
                this.propertyID = PropertyHelper.PropertyTagDic[propertyName].PropertyId;
                this.ValueType = (PropertyType)PropertyHelper.PropertyTagDic[propertyName].PropertyType;
                this.value = value;
            }
        }

        /// <summary>
        /// Initializes a new instance of the PropertyObj class.
        /// </summary>
        /// <param name="propertyId">Property ID.</param>
        /// <param name="propertyType">Property type code.</param>
        /// <param name="value">The value of property.</param>
        public PropertyObj(uint propertyId, ushort propertyType, object value)
        {
            this.propertyID = propertyId;
            this.ValueType = (PropertyType)propertyType;
            this.value = value;
        }

        /// <summary>
        /// Initializes a new instance of the PropertyObj class.
        /// </summary>
        /// <param name="propertyID">Property Code/ID</param>
        /// <param name="propertyType">Property type code</param>
        public PropertyObj(uint propertyID, uint propertyType)
        {
            this.propertyID = propertyID;
            this.ValueType = (PropertyType)propertyType;
            this.PropertyName = PropertyHelper.GetPropertyNameByID(propertyID, propertyType);
        }

        /// <summary>
        /// Gets property ID.
        /// </summary>
        public uint PropertyID
        {
            get { return (uint)this.propertyID; }
        }

        /// <summary>
        /// Gets or sets property name
        /// </summary>
        public PropertyNames PropertyName
        {
            get { return this.propertyName; }
            set { this.propertyName = value; }
        }

        /// <summary>
        /// Gets or sets value type.
        /// </summary>
        public PropertyType ValueType
        {
            get { return this.valueType; }
            set { this.valueType = value; }
        }

        /// <summary>
        /// Gets the id for type of value.
        /// </summary>
        public ushort ValueTypeCode
        {
            get { return (ushort)this.valueType; }
        }

        /// <summary> 
        /// Gets or sets value type.
        /// </summary>
        public object Value
        {
            get { return this.value; }
            set { this.value = value; }
        }
    }
}