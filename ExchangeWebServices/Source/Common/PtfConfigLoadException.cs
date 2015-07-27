//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.Serialization;

    /// <summary>
    /// The exception that is thrown when the ptfconfig file is loaded with failure.
    /// </summary>
    [Serializable]
    public class PtfConfigLoadException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the PtfConfigLoadException class.
        /// </summary>
        public PtfConfigLoadException()
            : base() 
        {
        }

        /// <summary>
        /// Initializes a new instance of the PtfConfigLoadException class with a specified error message.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        public PtfConfigLoadException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PtfConfigLoadException class with specified error message and exception.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="exception">The exception.</param>
        public PtfConfigLoadException(string message, Exception exception)
            : base(message, exception)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PtfConfigLoadException class with serialized data.
        /// </summary>
        /// <param name="info">The object that holds the serialized object data.</param>
        /// <param name="context">The contextual information about the source or destination.</param>
        protected PtfConfigLoadException(SerializationInfo info, StreamingContext context)
            : base(info, context) 
        {
        }
    }
}