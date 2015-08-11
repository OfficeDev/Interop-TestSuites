namespace Microsoft.Protocols.TestSuites.Common
{
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Xml;
    using System.Xml.Serialization;

    /// <summary>
    /// The ActiveSync request.
    /// </summary>
    /// <typeparam name="T">The generic type.</typeparam>
    public abstract class ActiveSyncRequestBase<T>
    {
        /// <summary>
        /// Gets or sets request data.
        /// </summary>
        public T RequestData { get; set; }

        /// <summary>
        /// Gets command parameters.
        /// </summary>
        public IDictionary<CmdParameterName, object> CommandParameters { get; private set; }

        /// <summary>
        /// Sets command parameters
        /// </summary>
        /// <param name="parameters">The parameters of the command</param>
        public void SetCommandParameters(IDictionary<CmdParameterName, object> parameters)
        {
            this.CommandParameters = parameters;
        }

        /// <summary>
        /// Get request data serialized xml.
        /// </summary>
        /// <returns>The result of serialized xml.</returns>
        public virtual string GetRequestDataSerializedXML()
        {
            if (null == this.RequestData)
            {
                return string.Empty;
            }

            string serializedXMLstring;

            MemoryStream ms = null;
            try
            {
                ms = new MemoryStream();
                using (XmlWriter stringWriter = new ActiveSyncXmlWriter(ms, Encoding.UTF8))
                {
                    XmlSerializer xmlSerializer = new XmlSerializer(this.RequestData.GetType());
                    xmlSerializer.Serialize(stringWriter, this.RequestData);
                    ms.Position = 0;
                    serializedXMLstring = new StreamReader(ms).ReadToEnd();
                }
            }
            finally
            {
                if (ms != null)
                {
                    ms.Dispose();
                }
            }

            return serializedXMLstring;
        }
    }
}