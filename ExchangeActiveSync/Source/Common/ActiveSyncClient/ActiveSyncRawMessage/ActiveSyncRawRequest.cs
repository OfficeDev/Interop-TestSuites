namespace Microsoft.Protocols.TestSuites.Common
{
    using System.Collections.Generic;

    /// <summary>
    /// Wrapper class contains MS-ASCMD command name, HTTP header and post body information
    /// </summary>
    public class ActiveSyncRawRequest
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the ActiveSyncRawRequest class
        /// </summary>
        /// <param name="parameters">The parameters of the command</param>
        /// <param name="requestBody">The request XML string</param>
        public ActiveSyncRawRequest(IDictionary<CmdParameterName, object> parameters, string requestBody)
        {
            this.HttpRequestBody = requestBody;
            this.CommandParameters = parameters;
        }

        /// <summary>
        /// Initializes a new instance of the ActiveSyncRawRequest class
        /// </summary>
        public ActiveSyncRawRequest()
        {
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the HTTP request body XML string
        /// </summary>
        public string HttpRequestBody { get; set; }

        /// <summary>
        /// Gets or sets the HTTP method Post or Option
        /// </summary>
        public string HttpMethod { get; set; }

        /// <summary>
        /// Gets or sets the MS-ASCMD command name
        /// </summary>
        public CommandName CommandName { get; set; }

        /// <summary>
        /// Gets the CommandParameters
        /// </summary>
        public IDictionary<CmdParameterName, object> CommandParameters { get; private set; }

        /// <summary>
        /// Gets or sets the content type of HTTP request
        /// </summary>
        public string ContentType { get; set; }
        #endregion
        
        /// <summary>
        /// Sets the CommandParameters
        /// </summary>
        /// <param name="parameters">The parameters of the command</param>
        public void SetCommandParameters(IDictionary<CmdParameterName, object> parameters)
        {
            this.CommandParameters = parameters;
        }
    }
}