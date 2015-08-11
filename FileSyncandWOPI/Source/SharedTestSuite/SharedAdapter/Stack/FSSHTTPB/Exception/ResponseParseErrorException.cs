namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;

    /// <summary>
    /// This class specifies the response parse error exception.
    /// </summary>
    [Serializable]
    public class ResponseParseErrorException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the ResponseParseErrorException class
        /// </summary>
        /// <param name="index">Specify the index of object</param>
        /// <param name="innerException">Specify the inner exception</param>
        public ResponseParseErrorException(int index, Exception innerException)
            : base(innerException.Message, innerException)
        {
            this.Index = index;
        }

        /// <summary>
        /// Initializes a new instance of the ResponseParseErrorException class
        /// </summary>
        /// <param name="index">Specify the index of object</param>
        /// <param name="message">Specify the exception message</param>
        /// <param name="innerException">Specify the inner exception</param>
        public ResponseParseErrorException(int index, string message, Exception innerException)
            : base(message, innerException)
        {
            this.Index = index;
        }

        /// <summary>
        /// Gets or sets index of object.
        /// </summary>
        public int Index { get; set; }
    }
}