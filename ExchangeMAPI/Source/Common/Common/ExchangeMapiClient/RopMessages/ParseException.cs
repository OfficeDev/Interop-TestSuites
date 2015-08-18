namespace Microsoft.Protocols.TestSuites.Common
{
    using System;

    /// <summary>
    /// Exception thrown in parsing process.
    /// </summary>
    public class ParseException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ParseException" /> class.
        /// </summary>
        /// <param name="message">The message of the exception</param>
        public ParseException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ParseException" /> class.
        /// </summary>
        public ParseException()
            : base()
        {
        }
    }
}