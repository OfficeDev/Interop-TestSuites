//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// Property structure.
    /// </summary>
    public class Property
    {
        /// <summary>
        /// Property name.
        /// </summary>
        private string name;

        /// <summary>
        /// Property type.
        /// </summary>
        private PropertyType type;

        /// <summary>
        /// Initializes a new instance of the <see cref="Property"/> class.
        /// </summary>
        /// <param name="propType">Property type.</param>
        public Property(PropertyType propType)
        {
            this.type = propType;
            this.name = string.Empty;
        }

        /// <summary>
        /// Gets or sets property name.
        /// </summary>
        public string Name
        {
            get
            {
                return this.name;
            }

            set
            {
                this.name = value;
            }
        }

        /// <summary>
        /// Gets or sets property type.
        /// </summary>
        public PropertyType Type
        {
            get
            {
                return this.type;
            }

            set
            {
                this.type = value;
            }
        }
    }
}