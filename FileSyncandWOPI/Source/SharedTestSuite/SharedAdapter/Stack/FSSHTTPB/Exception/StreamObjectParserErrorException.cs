namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;

    /// <summary>
    /// This class specifies the Stream object parse error exception.
    /// </summary>
    [Serializable]
    public class StreamObjectParseErrorException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the StreamObjectParseErrorException class
        /// </summary>
        /// <param name="index">Specify the index of object</param>
        /// <param name="streamObjectTypeName">Specify the stream type name</param>
        /// <param name="innerException">Specify the inner exception</param>
        public StreamObjectParseErrorException(int index, string streamObjectTypeName, Exception innerException)
            : base(innerException.Message, innerException)
        {
            this.Index = index;
            this.StreamObjectTypeName = streamObjectTypeName;
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectParseErrorException class
        /// </summary>
        /// <param name="index">Specify the index of object</param>
        /// <param name="streamObjectTypeName">Specify the stream type name</param>
        /// <param name="message">Specify the exception message</param>
        /// <param name="innerException">Specify the inner exception</param>
        public StreamObjectParseErrorException(int index, string streamObjectTypeName, string message, Exception innerException)
            : base(message, innerException)
        {
            this.Index = index;
            this.StreamObjectTypeName = streamObjectTypeName;
        }

        /// <summary>
        /// Gets or sets index of object.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Gets or sets stream object type name.
        /// </summary>
        public string StreamObjectTypeName { get; set; }
    }
}