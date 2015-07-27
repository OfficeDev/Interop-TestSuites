//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Text;

    /// <summary>
    /// An attribute indicate the encoding of a string field.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property
        | AttributeTargets.Field)]
    public sealed class StringEncodingAttribute : Attribute
    {
        /// <summary>
        /// The encoding type.
        /// </summary>
        private EncodingTypes encodingType;

        /// <summary>
        /// Initializes a new instance of the StringEncodingAttribute class.
        /// </summary>
        /// <param name="encodingType">Specify encoding type.</param>
        public StringEncodingAttribute(EncodingTypes encodingType)
        {
            this.encodingType = encodingType;
        }

        /// <summary>
        /// Encoding types. 
        /// </summary>
        public enum EncodingTypes
        { 
            /// <summary>
            /// ASCII encoding.
            /// </summary>
            ASCII = 0,

            /// <summary>
            /// Unicode encoding.
            /// </summary>
            Unicode = 1
        }

        /// <summary>
        /// Gets the encoding type.
        /// </summary>
        public EncodingTypes EncodingType
        {
            get { return this.encodingType; }
        }

        /// <summary>
        /// Gets corresponding Encoding
        /// </summary>
        public Encoding Encoding
        {
            get
            {
                switch (this.EncodingType)
                { 
                    case EncodingTypes.ASCII:
                        return Encoding.ASCII;
                    case EncodingTypes.Unicode:
                        return Encoding.Unicode;
                }

                return null;
            }
        }

        /// <summary>
        /// Get StringEncoding attribute from an object.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <returns>A StringEncoding instance have been specified to the object.</returns>
        public static StringEncodingAttribute GetStringEncoding(object obj)
        {
            return FieldHelper.GetFirstCustomAttribute<StringEncodingAttribute>(obj, false);
        }
    }
}