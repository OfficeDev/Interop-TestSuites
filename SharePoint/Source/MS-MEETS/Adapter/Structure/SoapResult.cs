//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This class aggregates the de-serialized soap response message, or the soap fault response message.
    /// </summary>
    /// <typeparam name="T">the type of the de-serialized soap response message.</typeparam>
    public class SoapResult<T>
    {
        /// <summary>
        /// Represents the de-serialized soap response message.
        /// </summary>
        private T result;

        /// <summary>
        /// Represents the soap fault response message.
        /// </summary>
        private SoapException exception;

        /// <summary>
        /// Initializes a new instance of the SoapResult class
        /// </summary>
        /// <param name="result">the de-serialized soap response message.</param>
        /// <param name="exception">the soap fault response message.</param>
        public SoapResult(T result, SoapException exception)
        {
            this.result = result;
            this.exception = exception;
        }

        /// <summary>
        /// Gets the de-serialized soap response message. 
        /// </summary>
        /// <value>the de-serialized soap response message.</value>
        public T Result
        {
            get
            {
                return this.result;
            }
        }

        /// <summary>
        /// Gets the soap fault response message.
        /// </summary>
        /// <value>the soap fault response message.</value>
        public SoapException Exception
        {
            get
            {
                return this.exception;
            }
        }

        /// <summary>
        /// Gets the SOAP fault code in the SOAP detail element
        /// </summary>
        /// <returns>The SOAP fault code value. Null if not exist</returns>
        public string GetErrorCode()
        {
            string errorCode = Common.ExtractErrorCodeFromSoapFault(this.exception);

            return errorCode;
        }
    }
}