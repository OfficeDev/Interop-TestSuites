//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System.ServiceModel.Channels;

    /// <summary>
    /// This class represents the factory for producing WOPI message encoder.
    /// </summary>
    public class WOPIMessageEncoderFactory : MessageEncoderFactory
    {
        /// <summary>
        /// Specify the message encoder.
        /// </summary>
        private MessageEncoder encoder;
        
        /// <summary>
        /// Specify the encoding type.
        /// </summary>
        private string encoding;
        
        /// <summary>
        /// Specify the message version.
        /// </summary>
        private MessageVersion messageVersion;

        /// <summary>
        /// Initializes a new instance of the WOPIMessageEncoderFactory class with the specified message version and char set version.
        /// </summary>
        /// <param name="messageVersion">Specify the message version.</param>
        /// <param name="encoding">Specify the encoding type.</param>
        public WOPIMessageEncoderFactory(MessageVersion messageVersion, string encoding)
        {
            this.messageVersion = messageVersion;
            this.encoding = encoding;
            this.encoder = new WOPIMessageEncoder(this);
        }

        /// <summary>
        /// Gets the char set version.
        /// </summary>
        public string CharSet
        {
            get
            {
                return this.encoding;
            }
        }

        /// <summary>
        /// Gets the message encoder.
        /// </summary>
        public override MessageEncoder Encoder
        {
            get { return this.encoder; }
        }

        /// <summary>
        /// Gets the message version.
        /// </summary>
        public override MessageVersion MessageVersion
        {
            get { return this.messageVersion; }
        }
    }
}