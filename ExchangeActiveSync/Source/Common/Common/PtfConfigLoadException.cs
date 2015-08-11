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