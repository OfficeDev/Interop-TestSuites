namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A connect request body for Mailbox Server Endpoint.
    /// </summary>
    public class ConnectRequestBody : MailboxRequestBodyBase
    {
        /// <summary>
        /// The value of rgbUserDN field in connect request body.
        /// </summary>
        private string rgbUserDN;

        /// <summary>
        /// Gets or sets the rgbUserDN field in connect request type request body.
        /// </summary>
        public string UserDN
        {
            get
            {
                return this.rgbUserDN;
            }

            set
            {
                if (value.EndsWith("\0") == false)
                {
                    value += "\0";
                }

                this.rgbUserDN = value;
            }
        }

        /// <summary>
        /// Gets or sets the ulFlags field in connect request type request body. 
        /// </summary>
        public uint Flags { get; set; }

        /// <summary>
        /// Gets or sets the ulCpid field in connect request type request body. 
        /// </summary>
        public uint Cpid { get; set; }

        /// <summary>
        /// Gets or sets the ulLcidSort field in connect request type request body. 
        /// </summary>
        public uint LcidSort { get; set; }

        /// <summary>
        /// Gets or sets the ulLcidString field in connect request type request body. 
        /// </summary>
        public uint LcidString { get; set; }

        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>The serialized data to be returned.</returns>
        public override byte[] Serialize()
        {
            List<byte> rawData = new List<byte>();

            rawData.AddRange(System.Text.Encoding.ASCII.GetBytes(this.UserDN));
            rawData.AddRange(BitConverter.GetBytes(this.Flags));
            rawData.AddRange(BitConverter.GetBytes(this.Cpid));
            rawData.AddRange(BitConverter.GetBytes(this.LcidSort));
            rawData.AddRange(BitConverter.GetBytes(this.LcidString));
            rawData.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            rawData.AddRange(this.AuxiliaryBuffer);

            return rawData.ToArray();
        }
    }
}