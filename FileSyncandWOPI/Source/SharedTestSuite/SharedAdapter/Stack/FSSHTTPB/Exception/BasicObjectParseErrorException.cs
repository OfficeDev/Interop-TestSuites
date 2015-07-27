//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;

    /// <summary>
    /// This class specifies the basic object parse error exception.
    /// </summary>
    [Serializable] 
    public class BasicObjectParseErrorException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the BasicObjectParseErrorException class
        /// </summary>
        /// <param name="index">Specify the index of object</param>
        /// <param name="basicTypeName">Specify the type name</param>
        /// <param name="innerException">Specify the inner exception</param>
        public BasicObjectParseErrorException(int index, string basicTypeName, Exception innerException)
            : base(innerException.Message, innerException)
        {
            this.Index = index;
            this.BasicObjectTypeName = basicTypeName;
        }

        /// <summary>
        /// Initializes a new instance of the BasicObjectParseErrorException class
        /// </summary>
        /// <param name="index">Specify the index of object</param>
        /// <param name="basicTypeName">Specify the basic type name</param>
        /// <param name="message">Specify the exception message</param>
        /// <param name="innerException">Specify the inner exception</param>
        public BasicObjectParseErrorException(int index, string basicTypeName, string message, Exception innerException)
            : base(message, innerException)
        {
            this.Index = index;
            this.BasicObjectTypeName = basicTypeName;
        }

        /// <summary>
        /// Gets or sets index of object.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Gets or sets basic object type name.
        /// </summary>
        public string BasicObjectTypeName { get; set; }
    }
}