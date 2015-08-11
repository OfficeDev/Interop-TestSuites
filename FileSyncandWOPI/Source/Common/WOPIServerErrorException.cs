namespace Microsoft.Protocols.TestSuites.Common
{
    using System;

    /// <summary>
    /// This class specifies the WOPI server error exception.
    /// </summary>
    public class WOPIServerErrorException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the WOPIServerErrorException class.
        /// </summary>
        public WOPIServerErrorException()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the WOPIServerErrorException class
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception</param>
        public WOPIServerErrorException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the WOPIServerErrorException class with a specified
        /// error message and a reference to the inner exception that is the cause of
        /// this exception.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="innerException">The exception that is the cause of the current exception. If the innerException
        /// parameter is not a null reference, the current exception is raised in a catch
        /// block that handles the inner exception.</param>
        public WOPIServerErrorException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}