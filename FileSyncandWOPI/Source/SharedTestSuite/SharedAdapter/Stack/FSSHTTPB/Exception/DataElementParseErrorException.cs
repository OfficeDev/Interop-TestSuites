namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;

    /// <summary>
    /// This class specifies the data element parse error exception.
    /// </summary>
    [Serializable]
    public class DataElementParseErrorException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the DataElementParseErrorException class
        /// </summary>
        /// <param name="index">Specify the index of object</param>
        /// <param name="innerException">Specify the inner exception</param>
        public DataElementParseErrorException(int index, Exception innerException)
            : base(innerException.Message, innerException)
        {
            this.Index = index;
        }

        /// <summary>
        /// Initializes a new instance of the DataElementParseErrorException class
        /// </summary>
        /// <param name="index">Specify the index of object</param>
        /// <param name="message">Specify the exception message</param>
        /// <param name="innerException">Specify the inner exception</param>
        public DataElementParseErrorException(int index, string message, Exception innerException)
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