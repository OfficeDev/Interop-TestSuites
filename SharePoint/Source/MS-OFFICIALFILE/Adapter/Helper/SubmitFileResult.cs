//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using System;

    /// <summary>
    /// The detailed data result for the SubmitFile WSDL operation.
    /// </summary>
    public class SubmitFileResult
    {
        /// <summary>
        /// This property indicates the result status code of the SubmitFile WSDL operation. 
        /// </summary>
        private SubmitFileResultCode? resultCode = null;

        /// <summary>
        /// This property indicates the URL of the file after the SubmitFile WSDL operation.
        /// </summary>
        private string resultUrl = null;

        /// <summary>
        /// This property indicates additional information returned by the SubmitFile operation.
        /// </summary>
        private string additionalInformation = null;

        /// <summary>
        /// This property indicates the result of custom file processing.
        /// </summary>
        private CustomProcessingResult customProcessingResult = null;

        /// <summary>
        /// Gets or sets the result status code of the SubmitFile WSDL operation. 
        /// </summary>
        public SubmitFileResultCode? ResultCode
        {
            get
            {
                return this.resultCode;
            }

            set
            {
                this.resultCode = value;
            }
        }

        /// <summary>
        /// Gets or sets the URL of the file after the SubmitFile WSDL operation.
        /// </summary>
        public string ResultUrl
        {
            get
            {
                return this.resultUrl;
            }

            set
            {
                this.resultUrl = value;
            }
        }

        /// <summary>
        /// Gets or sets additional information returned by the SubmitFile WSL operation. 
        /// </summary>
        public string AdditionalInformation
        {
            get
            {
                return this.additionalInformation;
            }

            set
            {
                this.additionalInformation = value;
            }
        }

        /// <summary>
        /// Gets or sets the result of custom file processing. 
        /// </summary>
        public CustomProcessingResult CustomProcessingResult
        {
            get
            {
                return this.customProcessingResult;
            }

            set
            {
                this.customProcessingResult = value;
            }
        }
    }
}
